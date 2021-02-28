Function Register-FileToDelete {
    <#
        .SYNOPSIS
            Registers a file/s or folder/s for deletion after a reboot.

        .DESCRIPTION
            Registers a file/s or folder/s for deletion after a reboot.

        .PARAMETER Source
            Collection of Files/Folders which will be marked for deletion after a reboot

        .NOTES
            Name: Register-FileToDelete
            Author: Boe Prox
            Created: 28 SEPT 2013

        .EXAMPLE
            Register-FileToDelete -Source 'C:\Users\Administrators\Desktop\Test.txt'
            True

            Description
            -----------
            Marks the file Test.txt for deletion after a reboot.

        .EXAMPLE
            Get-ChildItem -File -Filter *.txt | Register-FileToDelete -WhatIf
            What if: Performing operation "Mark for deletion" on Target "C:\Users\Administrator\Des
            ktop\SQLServerReport.ps1.txt".
            What if: Performing operation "Mark for deletion" on Target "C:\Users\Administrator\Des
            ktop\test.txt".


            Description
            -----------
            Uses a WhatIf switch to show what files would be marked for deletion.
    #>
    [cmdletbinding(
        SupportsShouldProcess = $True
    )]
    Param (
        [parameter(ValueFromPipeline=$True,
                  ValueFromPipelineByPropertyName=$True)]
        [Alias('FullName','File','Folder')]
        $Source = 'C:\users\Administrator\desktop\test.txt'    
    )
    Begin {
        Try {
            $null = [File]
        } Catch { 
            Write-Verbose 'Compiling code to create type'   
            Add-Type @"
            using System;
            using System.Collections.Generic;
            using System.Linq;
            using System.Text;
            using System.Runtime.InteropServices;
        
            public class Posh
            {
                public enum MoveFileFlags
                {
                    MOVEFILE_REPLACE_EXISTING           = 0x00000001,
                    MOVEFILE_COPY_ALLOWED               = 0x00000002,
                    MOVEFILE_DELAY_UNTIL_REBOOT         = 0x00000004,
                    MOVEFILE_WRITE_THROUGH              = 0x00000008,
                    MOVEFILE_CREATE_HARDLINK            = 0x00000010,
                    MOVEFILE_FAIL_IF_NOT_TRACKABLE      = 0x00000020
                }

                [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
                static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, MoveFileFlags dwFlags);
                public static bool MarkFileDelete (string sourcefile)
                {
                    bool brc = false;
                    brc = MoveFileEx(sourcefile, null, MoveFileFlags.MOVEFILE_DELAY_UNTIL_REBOOT);          
                    return brc;
                }
            }
"@
        }
    }
    Process {
        ForEach ($item in $Source) {
            Write-Verbose ('Attempting to resolve {0} to full path if not already' -f $item)
            $item = (Resolve-Path -Path $item).ProviderPath
            If ($PSCmdlet.ShouldProcess($item,'Mark for deletion')) {
                If (-NOT [Posh]::MarkFileDelete($item)) {
                    Try {
                        Throw (New-Object System.ComponentModel.Win32Exception)
                    } Catch {Write-Warning $_.Exception.Message}
                }
            }
        }
    }
}