# PowerShell Heartbeat Check

A small PowerShell script that keeps pinging a host and emails you when it goes down, then
emails you again when it comes back. It is meant for plain "is this box reachable?" monitoring
on Windows: no agent, no third-party modules, only built-in cmdlets.

**Languages:** English | [简体中文](README.zh-CN.md) 

## Features

- Continuous ICMP checks; interval and packets per sample are configurable
- Alert email after a configurable number of consecutive failures (default: 5)
- Recovery email once the target answers again
- Email delivery through a pending queue with exponential backoff (30s up to 900s), so an
  unreachable SMTP server causes no retry storm and no alert is lost
- Log file with automatic size-based rotation and retention
- SMTP port and TLS are configurable (defaults: 587 + TLS)
- A single failed check never terminates the monitor
- Emails include the monitoring host name, the first failure time, the resolved target
  addresses, sampling statistics and the most recent error
- Configurable log path; no external dependencies

## Requirements

- Windows with PowerShell 5.1 or PowerShell 7+
- An SMTP server reachable from the monitoring host
- Permission to send ICMP echo requests on the network

## Quick start

```powershell
.\powershell-heartbeat-check.ps1 `
    -Target 192.168.1.1 `
    -SmtpServer smtp.example.com `
    -From monitor@example.com `
    -To ops@example.com `
    -Credential (Get-Credential)
```

Plain-text internal relay on port 25:

```powershell
.\powershell-heartbeat-check.ps1 -Target 192.168.1.1 -SmtpServer 10.0.0.25 `
    -From monitor@example.com -To ops@example.com -SmtpPort 25 -UseSsl $false
```

Running unattended (scheduled task): store the credential once, then reuse it.

```powershell
Get-Credential | Export-Clixml C:\PingMonitor\smtp.cred

.\powershell-heartbeat-check.ps1 -Target 192.168.1.1 -SmtpServer smtp.example.com `
    -From monitor@example.com -To ops@example.com `
    -Credential (Import-Clixml C:\PingMonitor\smtp.cred)
```

## Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `-Target` | *required* | IP address or host name to ping |
| `-SmtpServer` | *required* | SMTP server address |
| `-From` | *required* | Sender email address |
| `-To` | *required* | Recipient address(es), comma separated |
| `-Credential` | none | SMTP credentials (optional) |
| `-LogPath` | `C:\PingMonitor\ping_log.txt` | Path of the log file |
| `-FailureThreshold` | `5` | Consecutive failures that trigger an alert |
| `-Interval` | `5` | Seconds between two checks |
| `-SmtpPort` | `587` | SMTP server port |
| `-UseSsl` | `$true` | Use TLS/SSL for the SMTP connection |
| `-PingCount` | `2` | ICMP packets per sample; any reply counts as success |
| `-PingTimeoutSeconds` | `3` | Timeout per ICMP packet (PowerShell 7+ only) |
| `-MaxMailRetrySeconds` | `900` | Upper bound of the email retry backoff |
| `-LogMaxSizeMB` | `10` | Size at which the log file is rotated |
| `-LogRetainCount` | `5` | Number of rotated log files to keep |

## Notes

- Log lines and email bodies are currently written in Chinese; code comments are in English.
- When run as a scheduled task, console output is discarded. The log file is the durable record.
- `Send-MailMessage` is marked obsolete by Microsoft and cannot use OAuth2. It is kept here so
  the script needs no external dependencies; moving to MailKit is the long-term option.
- The default `587 + TLS` suits modern SMTP services. For a plain-text internal relay, pass
  `-SmtpPort 25 -UseSsl $false`.
- The script runs until it is stopped (Ctrl+C, or stopping the scheduled task).

## License

MIT License, see [LICENSE](LICENSE).
