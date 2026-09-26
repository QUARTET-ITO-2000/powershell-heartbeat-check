<#
.SYNOPSIS
    持续性网络连通性检测脚本
.DESCRIPTION
    持续ping指定地址，当连续失败超过指定次数时发送邮件告警，并生成日志文件
.PARAMETER Target
    要ping的目标地址（IP或域名）
.PARAMETER SmtpServer
    SMTP服务器地址
.PARAMETER From
    发件人邮箱地址
.PARAMETER To
    收件人邮箱地址
.PARAMETER Credential
    SMTP认证凭据（可选）
.PARAMETER LogPath
    日志文件保存路径
.PARAMETER FailureThreshold
    连续失败次数阈值，默认5次
.PARAMETER Interval
    Ping间隔时间（秒），默认5秒
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Target,
    
    [Parameter(Mandatory=$true)]
    [string]$SmtpServer,
    
    [Parameter(Mandatory=$true)]
    [string]$From,
    
    [Parameter(Mandatory=$true)]
    [string]$To,
    
    [PSCredential]$Credential,
    
    [string]$LogPath = "C:\PingMonitor\ping_log.txt",
    
    [int]$FailureThreshold = 5,
    
    [int]$Interval = 5
)

# 创建日志目录
$logDir = Split-Path $LogPath -Parent
if (!(Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# 初始化变量
$failureCount = 0
$alertSent = $false
$startTime = Get-Date

# 日志函数
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # 输出到控制台
    Write-Host $logEntry
    
    # 写入日志文件
    Add-Content -Path $LogPath -Value $logEntry -Encoding UTF8
}

# 发送邮件函数
function Send-AlertEmail {
    param(
        [string]$Subject,
        [string]$Body
    )
    
    try {
        $mailParams = @{
            SmtpServer = $SmtpServer
            From = $From
            To = $To
            Subject = $Subject
            Body = $Body
            Encoding = [System.Text.Encoding]::UTF8
        }
        
        if ($Credential) {
            $mailParams.Credential = $Credential
        }
        
        Send-MailMessage @mailParams
        Write-Log "告警邮件发送成功" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "邮件发送失败: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# 显示启动信息
Write-Log "=== Ping监控脚本启动 ==="
Write-Log "监控目标: $Target"
Write-Log "日志文件: $LogPath"
Write-Log "失败阈值: $FailureThreshold 次"
Write-Log "检查间隔: $Interval 秒"
Write-Log "开始时间: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Log "="

try {
    # 主监控循环
    while ($true) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # 执行ping测试
        $pingResult = Test-Connection -ComputerName $Target -Count 1 -Quiet
        
        if ($pingResult) {
            # Ping成功
            if ($failureCount -gt 0) {
                Write-Log "连接恢复 - 目标: $Target, 之前连续失败: $failureCount 次" "SUCCESS"
                
                # 发送恢复通知
                if ($alertSent) {
                    $recoveryBody = @"
网络连接已恢复:

目标地址: $Target
恢复时间: $timestamp
中断时长: $((New-TimeSpan -Start $startTime -End (Get-Date)).ToString())
累计失败次数: $failureCount

此邮件为自动发送，请勿回复。
"@
                    Send-AlertEmail -Subject "网络恢复通知 - $Target" -Body $recoveryBody
                }
            }
            
            $failureCount = 0
            $alertSent = $false
            Write-Log "Ping成功 - 目标: $Target" "INFO"
        }
        else {
            # Ping失败
            $failureCount++
            Write-Log "Ping失败 - 目标: $Target, 连续失败: $failureCount 次" "WARNING"
            
            # 检查是否达到失败阈值
            if ($failureCount -ge $FailureThreshold -and !$alertSent) {
                Write-Log "达到失败阈值，发送告警邮件" "ALERT"
                
                $alertBody = @"
网络连通性告警:

目标地址: $Target
告警时间: $timestamp
连续失败次数: $failureCount
失败阈值: $FailureThreshold
检查间隔: $Interval 秒

建议立即检查网络连接状态。

此邮件为自动发送，请勿回复。
"@
                $emailSent = Send-AlertEmail -Subject "网络连通性告警 - $Target" -Body $alertBody
                if ($emailSent) {
                    $alertSent = $true
                }
            }
        }
        
        # 等待指定间隔
        Start-Sleep -Seconds $Interval
    }
}
catch {
    Write-Log "脚本执行出错: $($_.Exception.Message)" "ERROR"
    Write-Log "脚本终止" "INFO"
}
finally {
    Write-Log "=== Ping监控脚本停止 ==="
    Write-Log "运行时长: $((New-TimeSpan -Start $startTime -End (Get-Date)).ToString())"
}