<#
.SYNOPSIS
既知のフォルダーをまとめて作成します。
.DESCRIPTION
既知のフォルダー (Known Folders) をまとめて作成します。
一部のフォルダーは作成するのに管理者権限が必要です。
既知のフォルダーの一覧は Microsot のサイトにあります。
https://docs.microsoft.com/en-us/windows/desktop/shell/knownfolderid
.INPUTS
なし
.OUTPUTS
psobject[]
フォルダーのパス情報
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

[bool]$verbose = $VerbosePreference -eq "Continue"

[string]$source = @'
using System;
using System.Runtime.InteropServices;
using System.Security;

namespace Win32API {
	public class KnownFolder {
		public readonly HResult Result;
		public readonly string Path;
		
		public KnownFolder(string guidText, uint flags) {
			IntPtr pszPath = IntPtr.Zero;
			try {
				this.Result = (HResult)SHGetKnownFolderPath(new Guid(guidText), flags, IntPtr.Zero, out pszPath);
				this.Path = Marshal.PtrToStringAuto(pszPath);
			} finally {
				if (pszPath != IntPtr.Zero) Marshal.FreeCoTaskMem(pszPath);
			}
		}
		
		[DllImport("shell32.dll"), SuppressUnmanagedCodeSecurity]
		static extern int SHGetKnownFolderPath(
			[MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr pszPath);
	}
	
	public enum HResult {
		OK = 0,
		Fail = unchecked((int)0x80004005),
		NotFound = unchecked((int)0x80070002),
		AccessDenied = unchecked((int)0x80070005),
		InvalidArg = unchecked((int)0x80070057),
	}
}
'@
Add-Type -TypeDefinition $source -ErrorAction Stop

[uint32]$KF_FLAG_DEFAULT = 0x00000000
[uint32]$KF_FLAG_CREATE = 0x00008000

@(
	@{ guid="{5E6C858F-0E22-4760-9AFE-EA3317B67173}"; name="Profile" }
	@{ guid="{31C0DD25-9439-4F12-BF41-7FF4EDA38722}"; name="3D Objects" } # Win10から
	@{ guid="{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"; name="Desktop" }
	@{ guid="{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"; name="Personal" }
	@{ guid="{374DE290-123F-4565-9164-39C4925E467B}"; name="Downloads" }
	@{ guid="{4BD8D571-6D19-48D3-BE97-422220080E43}"; name="My Music" }
	@{ guid="{DE92C1C7-837F-4F69-A3BB-86E631204A23}"; name="Playlists" }
	@{ guid="{33E28130-4E1E-4676-835A-98395C3BC3BB}"; name="My Pictures" }
	@{ guid="{AB5FB87B-7CE2-4F83-915D-550846C9537B}"; name="Camera Roll" } # Win8.1から
	@{ guid="{3B193882-D3AD-4eab-965A-69829D1FB59F}"; name="SavedPictures" } # Win10から
	@{ guid="{b7bede81-df94-4682-a7d8-57a52620b86f}"; name="Screenshots" }  # Win8から
	@{ guid="{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"; name="PhotoAlbums" }
	@{ guid="{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"; name="My Video" }
	@{ guid="{EDC0FE71-98D8-4F4A-B920-C8DC133CB165}"; name="Captures" } # Win10から
	@{ guid="{7AD67899-66AF-43BA-9156-6AAD42E6C596}"; name="AppMods" } # Win10 1703から
	@{ guid="{56784854-C6CB-462b-8169-88E350ACB882}"; name="Contacts" }
	@{ guid="{1777F761-68AD-4D8A-87BD-30B759FA33DD}"; name="Favorites" }
	@{ guid="{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"; name="Links" }
	@{ guid="{2f8b40c2-83ed-48ee-b383-a1f157ec6f9a}"; name="Recorded Calls" } # Win10から
	@{ guid="{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"; name="SavedGames" }
	@{ guid="{7d1d3a04-debb-4115-95cf-2f29da2920da}"; name="Searches" }
	@{ guid="{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}"; name="OneDrive" } # Win8.1から
	@{ guid="{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}"; name="OneDriveDocuments" } # Win8.1から
	@{ guid="{C3F2459E-80D6-45DC-BFEF-1F769F2BE730}"; name="OneDriveMusic" } # Win8.1から
	@{ guid="{339719B5-8C47-4894-94C2-D8F77ADD44A6}"; name="OneDrivePictures" } # Win8.1から
	@{ guid="{767E6811-49CB-4273-87C2-20F355E1085B}"; name="OneDriveCameraRoll" } # Win8.1から
	@{ guid="{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"; name="AppData" }
	@{ guid="{915221FB-9EFE-4bda-8FD7-F78DCA774F87}"; name="CredentialManager" }
	@{ guid="{B88F4DAA-E7BD-49a9-B74D-02885A5DC765}"; name="CryptoKeys" }
	@{ guid="{10C07CD0-EF91-4567-B850-448B77CB37F9}"; name="DpapiKeys" }
	@{ guid="{54EED2E0-E7CA-4fdb-9148-0F4247291CFA}"; name="SystemCertificates" }
	@{ guid="{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"; name="Quick Launch" }
	@{ guid="{9e3995ab-1f9c-4f13-b827-48b24b6c7174}"; name="User Pinned" } # Win7から
	@{ guid="{BCB5256F-79F6-4CEE-B725-DC34E402FD46}"; name="ImplicitAppShortcuts" } # Win7から
	@{ guid="{008ca0b1-55b4-4c56-b8a8-4de4b299d3be}"; name="AccountPictures" } # Win8から
	@{ guid="{C5ABBF53-E17F-4121-8900-86626FC2C973}"; name="NetHood" }
	@{ guid="{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"; name="PrintHood" }
	@{ guid="{AE50C081-EBD2-438A-8655-8A092E34987A}"; name="Recent" }
	@{ guid="{8983036C-27C0-404B-8F08-102D10DCFD74}"; name="SendTo" }
	@{ guid="{A63293E8-664E-48DB-A079-DF759E0509F7}"; name="Templates" }
	@{ guid="{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}"; name="Libraries" } # Win7から
	@{ guid="{2B20DF75-1EDA-4039-8097-38798227D5B7}"; name="CameraRollLibrary" } # Win10から
	@{ guid="{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}"; name="DocumentsLibrary" } # Win7から
	@{ guid="{2112AB0A-C86A-4FFE-A368-0DE96E47012E}"; name="MusicLibrary" } # Win7から
	@{ guid="{A990AE9F-A03B-4E80-94BC-9912D7504104}"; name="PicturesLibrary" } # Win7から
	@{ guid="{E25B5812-BE88-4bd9-94B0-29233477B6C3}"; name="SavedPicturesLibrary" } # Win10から
	@{ guid="{491E922F-5643-4af4-A7EB-4E7A138D8174}"; name="VideosLibrary" } # Win7から
	@{ guid="{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"; name="Start Menu" }
	@{ guid="{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"; name="Programs" }
	@{ guid="{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"; name="Administrative Tools" }
	@{ guid="{B97D20BB-F46A-4C97-BA10-5E3608430854}"; name="Startup" }
	@{ guid="{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"; name="LocalAppData" }
	@{ guid="{A520A1A4-1780-4FF6-BD18-167343C5AF16}"; name="LocalAppDataLow" }
	@{ guid="{B2C5E279-7ADD-439F-B28C-C41FE1BBF672}"; name="AppDataDesktop" } # Win10 1709から
	@{ guid="{DBE8E08E-3053-4BBC-B183-2A7B2B191E59}"; name="Development Files" } # Win10から
	@{ guid="{7BE16610-1F7F-44AC-BFF0-83E15F2FFCA1}"; name="AppDataDocuments" } # Win10 1709から
	@{ guid="{7CFBEFBC-DE1F-45AA-B843-A542AC536CC9}"; name="AppDataFavorites" } # Win10 1709から
	@{ guid="{559D40A3-A036-40FA-AF61-84CB430A4D34}"; name="AppDataProgramData" } # Win10 1709から
	@{ guid="{A3918781-E5F2-4890-B3D9-A7E54332328C}"; name="Application Shortcuts" } # Win8から
	@{ guid="{9E52AB10-F80D-49DF-ACB8-4330F5687855}"; name="CDBurning" }
	@{ guid="{054FAE61-4DD8-4787-80B6-090220C4B700}"; name="GameTasks" }
	@{ guid="{C870044B-F49E-4126-A9C3-B52A1FF411E8}"; name="Ringtones" } # Win7から
	@{ guid="{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}"; name="Roamed Tile Images" } # Win8から
	@{ guid="{00BCFC5A-ED94-4e48-96A1-3F6217F21990}"; name="Roaming Tiles" } # Win8から
	@{ guid="{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}"; name="SearchHistoryFolder" } # Win8.1から
	@{ guid="{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}"; name="SearchTemplatesFolder" } # Win8.1から
	@{ guid="{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"; name="Gadgets" } # Win7まで
	@{ guid="{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"; name="Original Images" }
	@{ guid="{5cd7aee2-2219-4a67-b85d-6c9ce15660cb}"; name="UserProgramFiles" } # Win7から
	@{ guid="{bcbd3057-ca5c-4622-b42d-bc56db0ae516}"; name="UserProgramFilesCommon" } # Win7から
	@{ guid="{D9DC8A3B-B784-432E-A781-5A1130A75963}"; name="History" }
	@{ guid="{352481E8-33BE-4251-BA85-6007CAEDCF9D}"; name="Cache" }
	@{ guid="{2B0F765D-C0E9-4171-908E-08A611B84FF6}"; name="Cookies" }
	@{ guid="{DFDF76A2-C82A-4D63-906A-5644AC457385}"; name="Public" }
	@{ guid="{0482af6c-08f1-4c34-8c90-e17ec98b1e17}"; name="PublicAccountPictures" } # Win8から
	@{ guid="{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"; name="Common Desktop" }
	@{ guid="{ED4824AF-DCE4-45A8-81E2-FC7965083634}"; name="Common Documents" }
	@{ guid="{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"; name="CommonDownloads" }
	@{ guid="{48DAF80B-E6CF-4F4E-B800-0E69D84EE384}"; name="PublicLibraries" } # Win7から
	@{ guid="{1A6FDBA2-F42D-4358-A798-B74D745926C5}"; name="RecordedTVLibrary" } # Win7から
	@{ guid="{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"; name="CommonMusic" }
	@{ guid="{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"; name="SampleMusic" }
	@{ guid="{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"; name="SamplePlaylists" } # Win7まで
	@{ guid="{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"; name="CommonPictures" }
	@{ guid="{C4900540-2379-4C75-844B-64E6FAF8716B}"; name="SamplePictures" }
	@{ guid="{2400183A-6185-49FB-A2D8-4A392A602BA3}"; name="CommonVideo" }
	@{ guid="{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"; name="SampleVideos" }
	@{ guid="{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"; name="ProgramData" }
	@{ guid="{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"; name="OEM Links" }
	@{ guid="{5CE4A5E9-E4EB-479D-B89F-130C02886155}"; name="Device Metadata Store" } # Win7から
	@{ guid="{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"; name="PublicGameTasks" }
	@{ guid="{12D4C69E-24AD-4923-BE19-31321C43A767}"; name="RetailDemo" } # Win10から
	@{ guid="{E555AB60-153B-4D17-9F04-A5FE99FC15EC}"; name="CommonRingtones" }
	@{ guid="{B94237E7-57AC-4347-9151-B08C6C32D1F7}"; name="Common Templates" }
	@{ guid="{A4115719-D62E-491D-AA7C-E74B8BE3B067}"; name="Common Start Menu" }
	@{ guid="{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"; name="Common Programs" }
	@{ guid="{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"; name="Common Administrative Tools" }
	@{ guid="{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"; name="Common Startup" }
	@{ guid="{A440879F-87A0-4F7D-B700-0207B966194A}"; name="Common Start Menu Places" } # Win10から
	@{ guid="{F38BF404-1D43-42F2-9305-67DE0B28FC23}"; name="Windows" }
	@{ guid="{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"; name="Fonts" }
	@{ guid="{8AD10C31-2ADB-4296-A8F7-E4701232C972}"; name="ResourceDir" }
	@{ guid="{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"; name="LocalizedResourcesDir" }
	@{ guid="{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"; name="System" }
	@{ guid="{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"; name="SystemX86" }
	@{ guid="{0762D272-C50A-4BB0-A382-697DCD729B80}"; name="UserProfiles" }
	@{ guid="{905e63b6-c1bf-494e-b29c-65b732d3d21a}"; name="ProgramFiles" }
	@{ guid="{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"; name="ProgramFilesX86" }
	@{ guid="{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"; name="ProgramFilesCommon" }
	@{ guid="{DE974D24-D9C6-4D3E-BF91-F4455120B917}"; name="ProgramFilesCommonX86" }
	@{ guid="{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"; name="Default Gadgets" } # Win7まで
) |
	ForEach-Object {
		[string]$guid = $_["guid"]
		[string]$name = $_["name"]
		
		[string]$result = $null
		[Win32API.KnownFolder]$folder = New-Object Win32API.KnownFolder $guid, $KF_FLAG_DEFAULT
		if ($folder.Result -eq "OK") {
			if (!$verbose) { return }
			$result = $folder.Result
		} else {
			$folder = New-Object Win32API.KnownFolder $guid, $KF_FLAG_CREATE
			if ($folder.Result -eq "NotFound" -and !$verbose) { return }
			$result = if ($folder.Result -eq "OK") { "New" } else { $folder.Result }
		}
		
		# [pscustomobject]@{}が使えるのはPowerShell3.0から
		# Select-Objectはプロパティの表示順を指定するのに使っている
		Write-Output (
			New-Object psobject -Property @{
				"Name" = $name
				"Result" = $result
				"Path" = $folder.Path
			} | Select-Object Name, Result, Path
		)
	}
