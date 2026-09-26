<#
.SYNOPSIS
    持续性网络连通性检测脚本（连续失败达到阈值后发送邮件告警）

.DESCRIPTION
    持续 ping 指定目标，当连续失败次数达到阈值时发送邮件告警，目标恢复后发送恢复通知，
    并持续写入日志文件。

    本次修改相对旧版本的变化：
      1. 恢复邮件的“中断时长”改为按“首次失败时间”计算，不再误用脚本启动时间；
      2. 新增 -SmtpPort / -UseSsl（默认 587 + TLS），邮件发送改为指数退避重试：
         既不会每个循环都重试冲击 SMTP 服务器，也不会因为一次发信失败而永久丢失告警；
      3. 单次 ping 或检测过程出现异常不再终止脚本，主循环对异常免疫；
      4. 日志文件按体积自动轮转并保留固定份数，避免长期运行写满磁盘；
      5. 邮件正文补充监控主机名、首次失败时间、目标解析地址、采样统计与最近错误；
      6. 每次采样可发送多个 ICMP 包（默认 2 个），任一回包即视为本次采样成功，降低误报；
      7. 文件保存为 UTF-8（带 BOM），中文在 PowerShell 5.1 / 7 下均不会乱码；
      8. 日志写入与邮件正文显式使用“不带 BOM 的 UTF-8”，避免邮件正文出现 BOM 乱码。

.PARAMETER Target
    要 ping 的目标地址（IP 或域名）。
.PARAMETER SmtpServer
    SMTP 服务器地址。
.PARAMETER From
    发件人邮箱地址。
.PARAMETER To
    收件人邮箱地址（多个地址用英文逗号分隔）。
.PARAMETER Credential
    SMTP 认证凭据（可选）。无人值守场景建议用 Export-Clixml 加密保存后导入。
.PARAMETER LogPath
    日志文件保存路径，默认 C:\PingMonitor\ping_log.txt。
.PARAMETER FailureThreshold
    连续失败次数阈值，默认 5 次。
.PARAMETER Interval
    两次检测之间的等待时间（秒），默认 5 秒。
.PARAMETER SmtpPort
    SMTP 服务器端口，默认 587（邮件提交端口）。
.PARAMETER UseSsl
    是否使用 TLS/SSL 加密，默认 $true；内网明文中继（例如 25 端口）请显式传 -UseSsl $false。
.PARAMETER PingCount
    每次采样发送的 ICMP 包数量，默认 2 个，只要有一个包得到应答即视为本次采样成功。
.PARAMETER PingTimeoutSeconds
    单个 ICMP 包的超时时间（秒），默认 3；仅 PowerShell 7 及以上生效，低版本沿用系统默认超时。
.PARAMETER MaxMailRetrySeconds
    邮件重试退避的最大间隔（秒），默认 900。
.PARAMETER LogMaxSizeMB
    单个日志文件的最大体积（MB），超过后自动轮转，默认 10。
.PARAMETER LogRetainCount
    日志轮转后保留的历史文件数量，默认 5 份。

.EXAMPLE
    .\持续性ping且失败5次后发送邮件.ps1 -Target 192.168.1.1 -SmtpServer smtp.example.com `
        -From monitor@example.com -To ops@example.com -Credential (Get-Credential)
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Target,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SmtpServer,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$From,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$To,

    [PSCredential]$Credential,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath = 'C:\PingMonitor\ping_log.txt',

    [ValidateRange(1, 100000)]
    [int]$FailureThreshold = 5,

    [ValidateRange(1, 86400)]
    [int]$Interval = 5,

    [ValidateRange(1, 65535)]
    [int]$SmtpPort = 587,

    [bool]$UseSsl = $true,

    [ValidateRange(1, 10)]
    [int]$PingCount = 2,

    [ValidateRange(0, 60)]
    [int]$PingTimeoutSeconds = 3,

    [ValidateRange(10, 86400)]
    [int]$MaxMailRetrySeconds = 900,

    [ValidateRange(1, 10240)]
    [int]$LogMaxSizeMB = 10,

    [ValidateRange(1, 100)]
    [int]$LogRetainCount = 5
)

Set-StrictMode -Version Latest
# 监控脚本必须长期存活：默认不把非终止错误升级为终止错误，只在关键调用上单独使用 -ErrorAction Stop
$ErrorActionPreference = 'Continue'

# ==================== 全局状态 ====================
$startTime          = Get-Date
$monitorHost        = [System.Net.Dns]::GetHostName()
$script:utf8NoBom   = [System.Text.UTF8Encoding]::new($false)
$script:mailQueue   = [System.Collections.Generic.Queue[object]]::new()
$script:mailBackoffInit = 30  # 邮件重试退避起始值（秒），发送成功后回到该值
$script:mailBackoff     = $script:mailBackoffInit
$script:maxPending      = 10  # 待发通知队列上限，避免长期发不出邮件时无限堆积
$script:logBytes        = 0   # 自上次体积检查后写入的字节数
$script:logErrShown     = $false

# PowerShell 5.1 默认可能未启用 TLS 1.2，会导致连接现代 SMTP 服务器失败
try {
    $securityProtocol = [System.Net.ServicePointManager]::SecurityProtocol
    if ($securityProtocol.ToString() -notmatch 'Tls12') {
        [System.Net.ServicePointManager]::SecurityProtocol = $securityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    }
}
catch {
    # PowerShell 7 / .NET Core 下该设置已无实际作用，忽略即可
}

# ==================== 函数定义 ====================

function Write-Log {
    <#
    .SYNOPSIS
        同时输出到控制台和日志文件，并在需要时触发日志轮转。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter(Position = 1)]
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ALERT', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry  = "[$timestamp] [$Level] $Message"

    # 控制台输出：交互运行时可见；作为计划任务运行时该输出会被丢弃，日志文件才是可靠记录
    Write-Host $logEntry

    try {
        Test-LogRotation
        [System.IO.File]::AppendAllText($LogPath, $logEntry + [Environment]::NewLine, $script:utf8NoBom)
        $script:logBytes += $script:utf8NoBom.GetByteCount($logEntry) + 2
    }
    catch {
        if (-not $script:logErrShown) {
            $script:logErrShown = $true
            Write-Host "警告: 写入日志文件失败，后续写入错误将不再提示。原因: $($_.Exception.Message)"
        }
    }
}

function Test-LogRotation {
    <#
    .SYNOPSIS
        日志文件超过指定体积时改名归档，并清理超出保留份数的历史日志。
    #>
    [CmdletBinding()]
    param()

    # 每累计写入约 1MB 才真正访问一次文件系统，避免高频采样下的额外 IO
    if ($script:logBytes -lt 1MB) { return }
    $script:logBytes = 0

    if (-not (Test-Path -LiteralPath $LogPath)) { return }

    $logFile = Get-Item -LiteralPath $LogPath -ErrorAction SilentlyContinue
    if ($null -eq $logFile) { return }
    if ($logFile.Length -lt ($LogMaxSizeMB * 1MB)) { return }

    $logDir = Split-Path -LiteralPath $LogPath -Parent
    if ([string]::IsNullOrEmpty($logDir)) { $logDir = (Get-Location).Path }

    $baseName    = [System.IO.Path]::GetFileNameWithoutExtension($LogPath)
    $extension   = [System.IO.Path]::GetExtension($LogPath)
    $archivePath = Join-Path -Path $logDir -ChildPath ('{0}.{1:yyyyMMdd-HHmmss}{2}' -f $baseName, (Get-Date), $extension)

    try {
        Move-Item -LiteralPath $LogPath -Destination $archivePath -Force -ErrorAction Stop
    }
    catch {
        Write-Host "警告: 日志轮转失败，将继续写入原日志文件。原因: $($_.Exception.Message)"
        return
    }

    try {
        $pattern    = '{0}.*{1}' -f $baseName, $extension
        $staleFiles = Get-ChildItem -LiteralPath $logDir -Filter $pattern -File -ErrorAction Stop |
            Where-Object { $_.Name -ne [System.IO.Path]::GetFileName($LogPath) } |
            Sort-Object -Property LastWriteTime -Descending |
            Select-Object -Skip $LogRetainCount

        if ($staleFiles) {
            $staleFiles | Remove-Item -Force -ErrorAction Stop
        }
    }
    catch {
        Write-Host "警告: 清理历史日志失败。原因: $($_.Exception.Message)"
    }

    # 此时 $script:logBytes 已被重置，因此不会递归触发轮转
    Write-Log "日志已轮转: $archivePath（保留最近 $LogRetainCount 份）" 'INFO'
}

function Get-PingSample {
    <#
    .SYNOPSIS
        对目标做一次 ICMP 采样，返回成功与否、回包数量与平均延迟。
    .DESCRIPTION
        异常在函数内部被捕获，调用方不会因为单次采样失败而中断监控。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [int]$Count = 2,

        [int]$TimeoutSeconds = 0
    )

    $sample = [pscustomobject]@{
        Timestamp = Get-Date
        Success   = $false
        Sent      = $Count
        Replies   = 0
        AverageMs = $null
        Error     = $null
    }

    $pingParams = @{
        ComputerName = $ComputerName
        Count        = $Count
        ErrorAction  = 'SilentlyContinue'
    }
    # -TimeoutSeconds 是 PowerShell 7 才引入的参数，低版本必须省略，否则会报参数不存在
    if ($TimeoutSeconds -gt 0 -and $PSVersionTable.PSVersion.Major -ge 7) {
        $pingParams['TimeoutSeconds'] = $TimeoutSeconds
    }

    $pingErrors = @()
    $replies    = @(Test-Connection @pingParams -ErrorVariable pingErrors)

    $okReplies = [System.Collections.Generic.List[object]]::new()
    $latencies = [System.Collections.Generic.List[double]]::new()

    foreach ($reply in $replies) {
        if ($null -eq $reply) { continue }

        # PowerShell 5.1 返回 Win32_PingStatus 对象，StatusCode 非 0 表示本次请求失败
        $statusCode = $reply.PSObject.Properties['StatusCode']
        if ($statusCode -and ([int]($statusCode.Value) -ne 0)) { continue }

        # PowerShell 7 返回带 Status 属性的对象
        $statusText = $reply.PSObject.Properties['Status']
        if ($statusText -and ("$($statusText.Value)" -ne 'Success')) { continue }

        $okReplies.Add($reply)

        # 两个版本的延迟属性名不同，逐个尝试
        foreach ($latencyName in @('ResponseTime', 'Latency')) {
            $latencyProperty = $reply.PSObject.Properties[$latencyName]
            if ($latencyProperty -and $null -ne $latencyProperty.Value) {
                $latencies.Add([double]($latencyProperty.Value))
                break
            }
        }
    }

    $sample.Replies = $okReplies.Count
    $sample.Success = ($okReplies.Count -gt 0)

    if ($latencies.Count -gt 0) {
        $average          = ($latencies | Measure-Object -Average).Average
        $sample.AverageMs = [int]([Math]::Round($average, 0))
    }

    if (-not $sample.Success -and $pingErrors.Count -gt 0) {
        $firstError   = $pingErrors | Select-Object -First 1
        $errorMessage = $firstError.Exception.Message
        if ([string]::IsNullOrWhiteSpace($errorMessage)) {
            $errorMessage = $firstError.ToString()
        }
        # 合并换行，保证错误信息能塞进单行日志与邮件正文
        $sample.Error = ($errorMessage -replace '\s+', ' ').Trim()
    }

    return $sample
}

function Get-TargetAddressInfo {
    <#
    .SYNOPSIS
        解析目标地址，失败时返回可直接写入邮件正文的原因说明。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    try {
        $addresses = @(
            [System.Net.Dns]::GetHostAddresses($ComputerName) |
                ForEach-Object { $_.ToString() } |
                Sort-Object -Unique
        )
        if ($addresses.Count -gt 0) { return ($addresses -join ', ') }
        return '未解析到任何地址'
    }
    catch {
        return "解析失败（$($_.Exception.Message)）"
    }
}

function Format-Duration {
    <#
    .SYNOPSIS
        把 TimeSpan 格式化成“x天x小时x分x秒”，避免依赖 TimeSpan 的自定义格式串。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [TimeSpan]$Span
    )

    return ('{0}天{1}小时{2}分{3}秒' -f $Span.Days, $Span.Hours, $Span.Minutes, $Span.Seconds)
}

function Send-NotificationEmail {
    <#
    .SYNOPSIS
        发送一封通知邮件，返回是否成功（本函数不抛出异常）。
    .NOTES
        Send-MailMessage 已被微软标记为过时（不支持 OAuth2 等现代认证），
        后续可迁移到 MailKit / Send-MailKitMessage；这里保留它是为了脚本零外部依赖。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$Body
    )

    $mailParams = @{
        SmtpServer  = $SmtpServer
        Port        = $SmtpPort
        UseSsl      = $UseSsl
        From        = $From
        To          = $To
        Subject     = $Subject
        Body        = $Body
        Encoding    = $script:utf8NoBom
        ErrorAction = 'Stop'
    }

    if ($Credential) {
        $mailParams['Credential'] = $Credential
    }

    try {
        Send-MailMessage @mailParams
        return $true
    }
    catch {
        Write-Log "邮件发送失败: $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

function Add-Notification {
    <#
    .SYNOPSIS
        把一条通知放入待发队列。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('ALERT', 'RECOVERY')]
        [string]$Kind,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$Body
    )

    # 队列上限保护：长期发不出邮件时丢弃最旧的一条，避免无限堆积
    while ($script:mailQueue.Count -ge $script:maxPending) {
        $dropped = $script:mailQueue.Dequeue()
        Write-Log "待发通知过多，丢弃最旧的一条: [$($dropped.Kind)] $($dropped.Subject)" 'WARNING'
    }

    $script:mailQueue.Enqueue([pscustomobject]@{
        Kind        = $Kind
        Subject     = $Subject
        Body        = $Body
        Attempts    = 0
        NextAttempt = Get-Date
    })

    Write-Log "已加入待发通知队列: [$Kind] $Subject" 'INFO'
}

function Send-PendingNotifications {
    <#
    .SYNOPSIS
        尝试发送队首的待发通知；失败时按指数退避推迟下一次尝试，不阻塞主循环。
    #>
    [CmdletBinding()]
    param()

    if ($script:mailQueue.Count -eq 0) { return }

    $head = $script:mailQueue.Peek()
    if ((Get-Date) -lt $head.NextAttempt) { return }

    $head.Attempts++

    if (Send-NotificationEmail -Subject $head.Subject -Body $head.Body) {
        [void]$script:mailQueue.Dequeue()
        $script:mailBackoff = $script:mailBackoffInit
        Write-Log "通知已发送: $($head.Subject)（第 $($head.Attempts) 次尝试）" 'SUCCESS'
    }
    else {
        $head.NextAttempt   = (Get-Date).AddSeconds($script:mailBackoff)
        Write-Log "通知发送失败，$($script:mailBackoff) 秒后重试（累计尝试 $($head.Attempts) 次）" 'WARNING'
        $script:mailBackoff = [Math]::Min($script:mailBackoff * 2, $MaxMailRetrySeconds)
    }
}

# ==================== 启动信息 ====================
if ($Credential) { $credentialState = '已提供' } else { $credentialState = '未提供' }

Write-Log '=== Ping监控脚本启动 ==='
Write-Log "监控主机: $monitorHost"
Write-Log "监控目标: $Target"
Write-Log "日志文件: $LogPath"
Write-Log "失败阈值: $FailureThreshold 次"
Write-Log "检查间隔: $Interval 秒"
Write-Log "每次采样: $PingCount 个 ICMP 包"
Write-Log "SMTP: $SmtpServer 端口 $SmtpPort TLS:$UseSsl 凭据:$credentialState"
Write-Log "收件人: $To"
Write-Log "开始时间: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Log '='

# ==================== 主监控循环 ====================
$failureCount     = 0
$failureStartTime = $null      # 本轮故障的首次失败时间，用于计算真实中断时长
$alertRaised      = $false     # 本轮故障是否已经产生过告警（含仍在队列中的）
$lastError        = '无'

try {
    while ($true) {
        try {
            # 1) 先处理待发通知（含退避重试），保证 SMTP 恢复后告警能补发
            Send-PendingNotifications

            # 2) 采样
            $sample    = Get-PingSample -ComputerName $Target -Count $PingCount -TimeoutSeconds $PingTimeoutSeconds
            $timestamp = $sample.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')

            if ($sample.Success) {
                # ---------- 本次采样成功 ----------
                if ($failureCount -gt 0) {
                    $outage     = New-TimeSpan -Start $failureStartTime -End $sample.Timestamp
                    $outageText = Format-Duration -Span $outage

                    Write-Log "连接恢复 - 目标: $Target, 中断时长: $outageText, 累计失败: $failureCount 次, 当前延迟: $($sample.AverageMs) ms" 'SUCCESS'

                    if ($alertRaised) {
                        $recoveryBody = @"
网络连接已恢复

监控主机: $monitorHost
目标地址: $Target
恢复时间: $timestamp
首次失败时间: $($failureStartTime.ToString('yyyy-MM-dd HH:mm:ss'))
中断时长: $outageText
累计失败次数: $failureCount
当前延迟: $($sample.AverageMs) ms
最近一次错误: $lastError

此邮件为自动发送，请勿回复。
"@
                        Add-Notification -Kind 'RECOVERY' -Subject "网络恢复通知 - $Target" -Body $recoveryBody
                    }
                }

                $failureCount     = 0
                $failureStartTime = $null
                $alertRaised      = $false
                $lastError        = '无'

                Write-Log "Ping成功 - 目标: $Target, 延迟: $($sample.AverageMs) ms" 'INFO'
            }
            else {
                # ---------- 本次采样失败 ----------
                if ($failureCount -eq 0) {
                    $failureStartTime = $sample.Timestamp
                }
                $failureCount++

                if ($sample.Error) {
                    $lastError = $sample.Error
                }
                else {
                    $lastError = '请求超时或目标无应答'
                }

                Write-Log "Ping失败 - 目标: $Target, 连续失败: $failureCount 次, 原因: $lastError" 'WARNING'

                # 达到阈值时只入队一次，之后的失败不再重复告警，直到目标恢复
                if ($failureCount -ge $FailureThreshold -and -not $alertRaised) {
                    $outage       = New-TimeSpan -Start $failureStartTime -End $sample.Timestamp
                    $outageText   = Format-Duration -Span $outage
                    $resolvedInfo = Get-TargetAddressInfo -ComputerName $Target

                    $alertBody = @"
网络连通性告警

监控主机: $monitorHost
目标地址: $Target
目标解析地址: $resolvedInfo
告警时间: $timestamp
首次失败时间: $($failureStartTime.ToString('yyyy-MM-dd HH:mm:ss'))
已中断时长: $outageText
连续失败次数: $failureCount
失败阈值: $FailureThreshold
检查间隔: $Interval 秒
本次采样: 发送 $($sample.Sent) 个包, 收到 $($sample.Replies) 个
最近一次错误: $lastError

建议立即检查网络连接状态。

此邮件为自动发送，请勿回复。
"@

                    Add-Notification -Kind 'ALERT' -Subject "网络连通性告警 - $Target" -Body $alertBody
                    $alertRaised = $true
                }
            }
        }
        catch {
            # 单次检测出错不再终止监控
            Write-Log "本次检测发生异常，已跳过并继续监控: $($_.Exception.Message)" 'ERROR'
        }

        # 等待下一次采样
        Start-Sleep -Seconds $Interval
    }
}
finally {
    Write-Log '=== Ping监控脚本停止 ==='
    Write-Log "运行时长: $(Format-Duration -Span (New-TimeSpan -Start $startTime -End (Get-Date)))"
}
