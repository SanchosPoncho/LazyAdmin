#Windows Defender exclusions for Hyper-V
        Get-MpPreference
        $WinVer=(Get-WmiObject -class Win32_OperatingSystem).Caption
        If ($env:SystemDrive+"\ProgramData\Microsoft\Windows\Hyper-V\Snapshots"){
        Add-MpPreference -ExclusionPath $env:SystemDrive+"\ProgramData\Microsoft\Windows\Hyper-V\Snapshots"}
        Add-MpPreference -ExclusionPath (Get-VMHost).VirtualHardDiskPath
        Add-MpPreference -ExclusionPath (Get-VMHost).VirtualMachinePath
        Add-MpPreference -ExclusionProcess "vmms.exe"
        Add-MpPreference -ExclusionProcess "vmwp.exe"
        If($WinVer -imatch 2016){Add-MpPreference -ExclusionProcess "vmsp.exe"} #2016 only
        If($WinVer -imatch 2019){Add-MpPreference -ExclusionProcess "Vmcompute.exe"} #2019 only
        Add-MpPreference -ExclusionExtension ".vhd"
        Add-MpPreference -ExclusionExtension ".vhdx"
        Add-MpPreference -ExclusionExtension ".avhd"
        Add-MpPreference -ExclusionExtension ".avhdx"
        Add-MpPreference -ExclusionExtension ".vhds"
        Add-MpPreference -ExclusionExtension ".vhdpmem"
        Add-MpPreference -ExclusionExtension ".iso"
        Add-MpPreference -ExclusionExtension ".rct"
        Add-MpPreference -ExclusionExtension ".vsv"
        Add-MpPreference -ExclusionExtension ".bin"
        Add-MpPreference -ExclusionExtension ".vmcx"
        Add-MpPreference -ExclusionExtension ".vmrs"
        Add-MpPreference -ExclusionPath (Get-cluster).SharedVolumesRoot #"C:\ClusterStorage"
        Get-MpPreference