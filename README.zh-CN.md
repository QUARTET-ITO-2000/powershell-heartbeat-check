# PowerShell 心跳检测

一个持续 ping 目标主机的 PowerShell 小脚本：目标不通时发邮件告警，目标恢复后再发一封恢复通知。
适合在 Windows 上做"这台机器还活着吗"这类简单监控，不需要安装 agent，也不依赖任何第三方模块，
只用系统内置命令。

**语言：** [English](README.md) | 简体中文

## 功能

- 持续 ICMP 检测，检测间隔与每次发包数均可配置
- 连续失败达到阈值（默认 5 次）后发送告警邮件
- 目标恢复后发送恢复通知邮件
- 邮件通过待发队列 + 指数退避发送（30 秒起，上限 900 秒）：SMTP 不可用时不刷屏，也不会丢告警
- 日志文件按体积自动轮转，并保留固定份数
- SMTP 端口与 TLS 可配置（默认 587 + TLS）
- 单次检测失败不会终止监控
- 邮件包含监控主机名、首次失败时间、目标解析地址、采样统计与最近一次错误
- 日志路径可配置，无外部依赖

## 环境要求

- Windows + PowerShell 5.1 或 PowerShell 7+
- 监控主机可访问的 SMTP 服务器
- 网络允许发送 ICMP echo 请求

## 快速开始

```powershell
.\powershell-heartbeat-check.ps1 `
    -Target 192.168.1.1 `
    -SmtpServer smtp.example.com `
    -From monitor@example.com `
    -To ops@example.com `
    -Credential (Get-Credential)
```

内网 25 端口明文中继：

```powershell
.\powershell-heartbeat-check.ps1 -Target 192.168.1.1 -SmtpServer 10.0.0.25 `
    -From monitor@example.com -To ops@example.com -SmtpPort 25 -UseSsl $false
```

无人值守（计划任务）：先保存一次凭据，之后重复使用。

```powershell
Get-Credential | Export-Clixml C:\PingMonitor\smtp.cred

.\powershell-heartbeat-check.ps1 -Target 192.168.1.1 -SmtpServer smtp.example.com `
    -From monitor@example.com -To ops@example.com `
    -Credential (Import-Clixml C:\PingMonitor\smtp.cred)
```

## 参数

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `-Target` | *必填* | 要 ping 的目标 IP 或域名 |
| `-SmtpServer` | *必填* | SMTP 服务器地址 |
| `-From` | *必填* | 发件人邮箱地址 |
| `-To` | *必填* | 收件人邮箱地址，多个用逗号分隔 |
| `-Credential` | 无 | SMTP 认证凭据（可选） |
| `-LogPath` | `C:\PingMonitor\ping_log.txt` | 日志文件路径 |
| `-FailureThreshold` | `5` | 触发告警的连续失败次数 |
| `-Interval` | `5` | 两次检测之间的等待秒数 |
| `-SmtpPort` | `587` | SMTP 端口 |
| `-UseSsl` | `$true` | 是否使用 TLS/SSL |
| `-PingCount` | `2` | 每次采样发包数，任一回包即算成功 |
| `-PingTimeoutSeconds` | `3` | 单个 ICMP 包超时秒数（仅 PowerShell 7+ 生效） |
| `-MaxMailRetrySeconds` | `900` | 邮件重试退避的最大间隔秒数 |
| `-LogMaxSizeMB` | `10` | 日志文件轮转阈值（MB） |
| `-LogRetainCount` | `5` | 轮转后保留的历史日志份数 |

## 注意事项

- 日志与邮件正文文案目前为中文，代码注释为英文。
- 作为计划任务运行时控制台输出会被丢弃，日志文件才是可靠记录。
- `Send-MailMessage` 已被微软标记为过时且不支持 OAuth2，保留它是为了让脚本零外部依赖；
  长期建议迁移到 MailKit。
- 默认 `587 + TLS` 适用于现代 SMTP 服务；内网明文中继请传 `-SmtpPort 25 -UseSsl $false`。
- 脚本会持续运行，直到 Ctrl+C 或停止计划任务。

## 许可

MIT 许可协议，详见 [LICENSE](LICENSE)。
