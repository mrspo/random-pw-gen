Function New-RandomPassword
{
    <#
    .SYNOPSIS
        Generate a random password.

    .DESCRIPTION
        This function generates a random password with a specified length and character set. 
        The password length can be between 12 and 40 characters. The character set can be 'Friendly', 'All', or 'None'.
        'Friendly' character set includes lowercase, uppercase, numeric, and special characters except ?\/*%#&.
        'All' character set includes full set of special characters.
        'None' character set includes only lowercase, uppercase, and numeric characters.

    .PARAMETER PasswordLength
        An integer that specifies the length of the password. The default value is 20.

    .PARAMETER specialChar
        A string that specifies the character set to use. The default value is 'Friendly'. 
        The options are 'Friendly', 'All', and 'None'.

    .EXAMPLE
        New-RandomPassword

        This command generates a 20-character password using friendly set of special characters.

    .EXAMPLE
        New-RandomPassword -PasswordLength 15 -specialChar 'All'

        This command generates a 15-character password using full set of special characters.

    .NOTES

        AUTHOR  : Kyle Espinoza
        LASTEDIT: 02/28/2024 10:00:00
        VERSION : 1.0
        KEYWORDS: PW, PASSWORD, RANDOM, GENERATE, SECURITY, CREDENTIALS, CREDS, NEW
    #>
    
    param(
        [Parameter(Mandatory=$false)]
        [ValidateRange(12,40)]
        [int]$PasswordLength = 20,

        [Parameter(Mandatory=$false)]
        [ValidateSet('Friendly','All','None')]
        [string]$specialChar = 'Friendly'
        )
 
    #ASCII Character set for Password
    switch ($specialChar)
    {
        Friendly
        {
            $CharacterSet = @{
                Lowercase   = (97..122) | Get-Random -Count 10 | ForEach-Object {[char]$_}
                Uppercase   = (65..90)  | Get-Random -Count 10 | ForEach-Object {[char]$_}
                Numeric     = (48..57)  | Get-Random -Count 10 | ForEach-Object {[char]$_}
                SpecialChar = '-!()*,.:;[]^_{|}~+<=>'.ToCharArray() | Get-Random -Count 10 | ForEach-Object {[char]$_}
            }
        }

        All
        {
            $CharacterSet = @{
                Lowercase   = (97..122) | Get-Random -Count 10 | ForEach-Object {[char]$_}
                Uppercase   = (65..90)  | Get-Random -Count 10 | ForEach-Object {[char]$_}
                Numeric     = (48..57)  | Get-Random -Count 10 | ForEach-Object {[char]$_}
                SpecialChar = (33..47)+(58..64)+(91..96)+(123..126) | Get-Random -Count 10 | ForEach-Object {[char]$_}
            }
        }

        None
        {
            $CharacterSet = @{
                Lowercase   = (97..122) | Get-Random -Count 10 | ForEach-Object {[char]$_}
                Uppercase   = (65..90)  | Get-Random -Count 10 | ForEach-Object {[char]$_}
                Numeric     = (48..57)  | Get-Random -Count 10 | ForEach-Object {[char]$_}
                SpecialChar = @()
            }
        }
    }

    #Frame Random Password from given character set
    $StringSet = $CharacterSet.Uppercase + $CharacterSet.Lowercase + $CharacterSet.Numeric + $CharacterSet.SpecialChar
    
    -join(Get-Random -Count $PasswordLength -InputObject $StringSet)
}