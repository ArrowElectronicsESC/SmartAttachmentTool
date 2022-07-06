param(
    [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.String]
    $ArrowName
)

Connect-ExchangeOnline -ShowProgress $false -ShowBanner:$False

$user = Get-EXORecipient -Identity $ArrowName -ErrorAction SilentlyContinue

$user.PrimarySmtpAddress