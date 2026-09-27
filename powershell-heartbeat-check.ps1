<#
.SYNOPSIS
    Continuous network connectivity monitor that sends an email alert after a
    configurable number of consecutive failures.

.DESCRIPTION
    Continuously pings the target. When the number of consecutive failures reaches the
    threshold an alert email is sent; once the target answers again a recovery notification
    is sent. Every event is also written to a log file.

    Changes compared with the previous version:
      1. The outage duration in the recovery email is now measured from the first failure
         time instead of the script start time (which produced meaningless values);
      2. New -SmtpPort / -UseSsl parameters (default 587 + TLS) and email delivery through a
         queue with exponential backoff: the SMTP server is no longer hit on every loop, and
         a single delivery failure no longer loses the alert permanently;
      3. A single ping or check failure no longer terminates the script; the main loop
         tolerates errors;
      4. The log file rotates by size and keeps a fixed number of archives, so a long run
         cannot fill the disk;
      5. Alert emails now include the monitoring host name, the first failure time, the
         resolved target addresses, sampling statistics and the most recent error;
      6. Each sample can send several ICMP packets (2 by default); any reply counts as a
         successful sample, which reduces false positives;
      7. The file is stored as UTF-8 (with BOM) so Chinese text is not garbled on
         PowerShell 5.1 / 7;
      8. Log writes and email bodies explicitly use UTF-8 without BOM to avoid BOM artefacts
         in the message body.

.PARAMETER Target
    Target address to ping (IP address or host name).
.PARAMETER SmtpServer
    SMTP server address.
.PARAMETER From
    Sender email address.
.PARAMETER To
    Recipient email address (separate multiple addresses with commas).
.PARAMETER Credential
    SMTP credentials (optional). For unattended runs, save them encrypted with Export-Clixml
    and import them again.
.PARAMETER LogPath
    Path of the log file. Default: C:\PingMonitor\ping_log.txt
.PARAMETER FailureThreshold
    Number of consecutive failures that triggers an alert. Default: 5
.PARAMETER Interval
    Seconds to wait between two checks. Default: 5
.PARAMETER SmtpPort
    SMTP server port. Default: 587 (mail submission port)
.PARAMETER UseSsl
    Whether to use TLS/SSL. Default: $true. For a plain-text internal relay (for example
    port 25) pass -UseSsl $false explicitly.
.PARAMETER PingCount
    Number of ICMP packets per sample. Default: 2. A single reply is enough to treat the
    sample as successful.
.PARAMETER PingTimeoutSeconds
    Timeout of a single ICMP packet in seconds. Default: 3. Honoured on PowerShell 7 and
    later only; older versions keep the system default timeout.
.PARAMETER MaxMailRetrySeconds
    Upper bound of the email retry backoff in seconds. Default: 900
.PARAMETER LogMaxSizeMB
    Maximum size of a single log file in MB before it is rotated. Default: 10
.PARAMETER LogRetainCount
    Number of rotated log files to keep. Default: 5

.EXAMPLE
    .\powershell-heartbeat-check.ps1 -Target 192.168.1.1 -SmtpServer smtp.example.com `
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
# A monitoring script has to survive for a long time, so non-terminating errors are not
# promoted to terminating ones globally; -ErrorAction Stop is used only where it is needed.
$ErrorActionPreference = 'Continue'

# ==================== Global state ====================
$startTime          = Get-Date
$monitorHost        = [System.Net.Dns]::GetHostName()
$script:utf8NoBom   = [System.Text.UTF8Encoding]::new($false)
$script:mailQueue   = [System.Collections.Generic.Queue[object]]::new()
$script:mailBackoffInit = 30  # Initial email retry backoff in seconds; restored after a successful send
$script:mailBackoff     = $script:mailBackoffInit
$script:maxPending      = 10  # Pending-notification queue cap, so it cannot grow while email is down
$script:logBytes        = 0   # Bytes written since the last size check
$script:logErrShown     = $false

# PowerShell 5.1 may not enable TLS 1.2 by default, which breaks modern SMTP servers
try {
    $securityProtocol = [System.Net.ServicePointManager]::SecurityProtocol
    if ($securityProtocol.ToString() -notmatch 'Tls12') {
        [System.Net.ServicePointManager]::SecurityProtocol = $securityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    }
}
catch {
    # This setting has no effect on PowerShell 7 / .NET Core, so any failure here is ignored
}

# ==================== Function definitions ====================

function Write-Log {
    <#
    .SYNOPSIS
        Writes one line to the console and to the log file, rotating the log when needed.
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

    # Console output is visible in interactive runs only; a scheduled task discards it,
    # so the log file is the durable record
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
        Renames the log file to an archive once it exceeds the configured size and removes
        archives beyond the retention count.
    #>
    [CmdletBinding()]
    param()

    # Only touch the file system after roughly 1MB has been written, to avoid extra IO
    # on frequent samples
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

    # $script:logBytes was reset above, so this cannot trigger a recursive rotation
    Write-Log "日志已轮转: $archivePath（保留最近 $LogRetainCount 份）" 'INFO'
}

function Get-PingSample {
    <#
    .SYNOPSIS
        Takes one ICMP sample of the target and returns success, reply count and average latency.
    .DESCRIPTION
        Exceptions are caught inside the function, so a failed sample never interrupts
        the caller.
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
    # -TimeoutSeconds only exists on PowerShell 7 and later; older versions must omit it
    # or the call fails with an unknown parameter error
    if ($TimeoutSeconds -gt 0 -and $PSVersionTable.PSVersion.Major -ge 7) {
        $pingParams['TimeoutSeconds'] = $TimeoutSeconds
    }

    $pingErrors = @()
    $replies    = @(Test-Connection @pingParams -ErrorVariable pingErrors)

    $okReplies = [System.Collections.Generic.List[object]]::new()
    $latencies = [System.Collections.Generic.List[double]]::new()

    foreach ($reply in $replies) {
        if ($null -eq $reply) { continue }

        # PowerShell 5.1 returns Win32_PingStatus objects; a non-zero StatusCode means the
        # request failed
        $statusCode = $reply.PSObject.Properties['StatusCode']
        if ($statusCode -and ([int]($statusCode.Value) -ne 0)) { continue }

        # PowerShell 7 returns objects with a Status property
        $statusText = $reply.PSObject.Properties['Status']
        if ($statusText -and ("$($statusText.Value)" -ne 'Success')) { continue }

        $okReplies.Add($reply)

        # The latency property is named differently in each version, so try both names
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
        # Collapse line breaks so the message fits into a single log line and email body
        $sample.Error = ($errorMessage -replace '\s+', ' ').Trim()
    }

    return $sample
}

function Get-TargetAddressInfo {
    <#
    .SYNOPSIS
        Resolves the target address; on failure it returns a reason string that can be used
        directly in the email body.
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
        Formats a TimeSpan as "XdXhXmXs" without relying on TimeSpan custom format strings.
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
        Sends one notification email and returns whether it succeeded (never throws).
    .NOTES
        Send-MailMessage is marked as obsolete by Microsoft (it cannot do OAuth2 or other
        modern authentication). Migrating to MailKit / Send-MailKitMessage is an option;
        it is kept here so the script needs no external dependencies.
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
        Adds one notification to the pending queue.
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

    # Queue cap: drop the oldest entry when email has been unavailable for a long time so
    # the queue cannot grow without limit
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
        Tries to send the notification at the head of the queue. On failure the next attempt
        is deferred with exponential backoff, without blocking the main loop.
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

# ==================== Startup information ====================
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

# ==================== Main monitoring loop ====================
$failureCount     = 0
$failureStartTime = $null      # First failure of the current outage, used to compute the real outage duration
$alertRaised      = $false     # Whether an alert was already raised for this outage (including a queued one)
$lastError        = '无'

try {
    while ($true) {
        try {
            # 1) Drain pending notifications first (with backoff retries) so queued alerts
            #    are delivered once SMTP recovers
            Send-PendingNotifications

            # 2) Take a sample
            $sample    = Get-PingSample -ComputerName $Target -Count $PingCount -TimeoutSeconds $PingTimeoutSeconds
            $timestamp = $sample.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')

            if ($sample.Success) {
                # ---------- sample succeeded ----------
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
                # ---------- sample failed ----------
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

                # Enqueue the alert only once when the threshold is crossed; later failures
                # stay silent until the target recovers
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
            # A failed check no longer terminates the monitor
            Write-Log "本次检测发生异常，已跳过并继续监控: $($_.Exception.Message)" 'ERROR'
        }

        # Wait for the next sample
        Start-Sleep -Seconds $Interval
    }
}
finally {
    Write-Log '=== Ping监控脚本停止 ==='
    Write-Log "运行时长: $(Format-Duration -Span (New-TimeSpan -Start $startTime -End (Get-Date)))"
}
