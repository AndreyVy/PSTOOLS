function New-Password {
	PARAM (
		[Int32]$Length = 12,
		[Int32]$NumberOfNonAlphanumericCharacters = 5,
		[Int32]$Count = 1
	)
	BEGIN { Add-Type -AssemblyName System.web }

	PROCESS {
		1..$Count | ForEach-Object {
			[System.Web.Security.Membership]::GeneratePassword($Length, $NumberOfNonAlphanumericCharacters)
		}
	}
}

function New-Password {
    param ([int]$Length = 12)
    $lowCharSet = [char[]]([char]33..[char]95)
    $UpCharSet = [char[]]([char]97..[char]126)
    $Digits = 0..9

    do {
        $Pass = (($lowCharSet + $UpCharSet + $Digits) | Sort-Object { Get-Random })[0..$Length] -join ''
        if (($Pass -cmatch '[a-z]') -and ($Pass -cmatch '[A-Z]') `
        -and ($Pass -cmatch '[0-9]') -and ($Pass -match '[^a-zA-Z0-9]')){
            return $Pass
        } else { continue }
    } while ($true)
}

function New-Password {
	[CmdletBinding()]
	PARAM (
		[ValidateNotNull()]
		[int]$Length = 12,
        [ValidateRange(1,256)]
        [Int]$Count = 1
	)#PARAM

	BEGIN {
		# Create ScriptBlock with the ASCII Char Codes
		$PasswordCharCodes = { 33..126 }.invoke()


        # Exclude some ASCII Char Codes from the ScriptBlock
        #  Excluded characters are ",',.,/,1,<,>,`,O,0,l,|
		#  See http://www.asciitable.com/ for mapping
		34, 39, 46, 47, 49, 60, 62, 96, 48, 79, 108, 124 | ForEach-Object { [void]$PasswordCharCodes.Remove($_) }
		$PasswordChars = [char[]]$PasswordCharCodes
	}#BEGIN

	PROCESS {
        1..$count | ForEach-Object {
            # Password of 4 characters or longer
		    IF ($Length -gt 4)
		    {

			    DO
			    {
				    # Generate a Password of the length requested
				    $NewPassWord = $(foreach ($i in 1..$length) { Get-Random -InputObject $PassWordChars }) -join ''
			    }#Do
			    UNTIL (
			    # Make sure it contains an Upercase and Lowercase letter, a number and another special character
			    ($NewPassword -cmatch '[A-Z]') -and
			    ($NewPassWord -cmatch '[a-z]') -and
			    ($NewPassWord -imatch '[0-9]') -and
			    ($NewPassWord -imatch '[^A-Z0-9]')
			    )#Until
		    }#IF
            # Password Smaller than 4 characters
		    ELSE
		    {
			    $NewPassWord = $(foreach ($i in 1..$length) { Get-Random -InputObject $PassWordChars }) -join ''
		    }#ELSE

		    # Output a new password
		    Write-Output $NewPassword
        }
	} #PROCESS
	END {
        # Cleanup
		Remove-Variable -Name NewPassWord -ErrorAction 'SilentlyContinue'
	} #END
} #Function