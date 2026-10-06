$ErrorActionPreference = 'Stop'
try {
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        throw 'Install GitHub CLI for Windows from https://cli.github.com/, then reopen this setup.'
    }
    foreach ($tokenName in @('GH_TOKEN', 'GITHUB_TOKEN', 'GH_ENTERPRISE_TOKEN', 'GITHUB_ENTERPRISE_TOKEN')) {
        if ([Environment]::GetEnvironmentVariable($tokenName)) {
            throw 'A token environment override is present. Remove it before using secure browser login.'
        }
    }
    Write-Host 'Sign in through GitHub in your browser. This setup does not publish any reports.'
    & gh auth login --hostname github.com --git-protocol https --web
    if ($LASTEXITCODE -ne 0) { throw 'GitHub browser sign-in did not complete.' }
    $authJson = & gh auth status --hostname github.com --active --json hosts
    if ($LASTEXITCODE -ne 0) { throw 'Unable to check GitHub authentication.' }
    $authState = $authJson | ConvertFrom-Json
    $accounts = @($authState.hosts.'github.com')
    if ($accounts.Count -ne 1 -or $accounts[0].state -ne 'success' -or
        $accounts[0].active -ne $true -or $accounts[0].tokenSource -ne 'keyring') {
        throw 'Secure credential-store login was not verified. Do not activate publishing. Plaintext fallback is not accepted.'
    }
    Write-Host 'Secure GitHub authentication verified. Publishing has NOT been activated.'
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
