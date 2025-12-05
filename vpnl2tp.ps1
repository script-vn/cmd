# Ví dụ cấu trúc đúng với GitHub raw
$uri = 'http://script-vn.github.io/cmd/vpn-l2tp-ios.ps1'
# Chỉ khi $uri trả về đúng text PowerShell thì mới iex
Invoke-RestMethod $uri | Invoke-Expression
