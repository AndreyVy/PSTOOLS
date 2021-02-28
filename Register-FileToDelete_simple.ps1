Add-Type @"
					using System;
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

			[Posh]::MarkFileDelete($FullFilePath)