# Network Connectivity and Speed Monitor

A PowerShell script for continuous monitoring of network connectivity and performance, designed for network administrators and IT professionals who need to track internet reliability and speed over time.

## Features

- Continuous network connectivity monitoring with multiple target servers
- Regular speed testing using standardized test files 
- Detailed CSV logging with hourly rotation
- System event logging integration
- Heartbeat monitoring system
- Prevention of system sleep during monitoring
- Automatic log directory creation and management
- Error handling and recovery

## Requirements

- Windows PowerShell 5.1 or later
- Administrator privileges (required for Event Log creation)
- Internet connection

## Configuration

Default intervals:
- Connectivity tests: Every 30 seconds
- Speed tests: Every 5 minutes  
- Heartbeat updates: Every 30 seconds

Test targets:
- Connectivity: Google DNS (8.8.8.8), Cloudflare DNS (1.1.1.1), OpenDNS (208.67.222.222)
- Speed: 1MB test files from reliable public sources

## Output

The script creates a `NetworkLogs` directory containing:
- Connectivity logs (CSV): Response times and success rates
- Speed test logs (CSV): Download speeds and test completion status
- Heartbeat file (JSON): Current script status and last test timestamps

## Usage

### Run as Administrator
```.\NetworkMonitor.ps1```

###Log Format
Connectivity Logs

Timestamp,SuccessfulPings,TotalTargets,AverageResponseTime,Status
2024-11-20 10:00:00,3,3,45.5,Connected
Speed Test Logs
Timestamp,DownloadSpeed,TestsCompleted,TotalTests
2024-11-20 10:05:00,50.25,2,2

### Error Handling
- Automatic recovery from network failures
- Event Log integration for error tracking
- Continuous operation despite individual test failures
  
## Known Limitations
- Requires administrative privileges for initial setup
- Speed tests use public test files which may have varying availability
- Download-only speed testing (no upload speed measurements)
- 
## Future Improvements
- Upload speed testing capability
- Custom test target configuration
- Configuration file support


## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

License
MIT License

## Acknowledgments
- Test files provided by various public sources
- PowerShell community for system interaction methods
