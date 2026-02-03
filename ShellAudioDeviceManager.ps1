# AudioDeviceManagerInteractive.ps1
# Interactive Audio Device Manager (Integrated Mic & Audio Toggle Edition)

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace AudioSwitcher {
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDevice {
        [PreserveSig] int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        [PreserveSig] int OpenPropertyStore(int stgmAccess, out IntPtr ppProperties);
        [PreserveSig] int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
        [PreserveSig] int GetState(out int pdwState);
    }
    
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceEnumerator {
        [PreserveSig] int EnumAudioEndpoints(int dataFlow, int dwStateMask, out IntPtr ppDevices);
        [PreserveSig] int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppEndpoint);
    }
    
    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioEndpointVolume {
        [PreserveSig] int RegisterControlChangeNotify(IntPtr pNotify);
        [PreserveSig] int UnregisterControlChangeNotify(IntPtr pNotify);
        [PreserveSig] int GetChannelCount(out int pnChannelCount);
        [PreserveSig] int SetMasterVolumeLevel(float fLevelDB, ref Guid pguidEventContext);
        [PreserveSig] int SetMasterVolumeLevelScalar(float fLevel, ref Guid pguidEventContext);
        [PreserveSig] int GetMasterVolumeLevel(out float pfLevelDB);
        [PreserveSig] int GetMasterVolumeLevelScalar(out float pfLevel);
        [PreserveSig] int SetChannelVolumeLevel(uint nChannel, float fLevelDB, ref Guid pguidEventContext);
        [PreserveSig] int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, ref Guid pguidEventContext);
        [PreserveSig] int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
        [PreserveSig] int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
        [PreserveSig] int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, ref Guid pguidEventContext);
        [PreserveSig] int GetMute([MarshalAs(UnmanagedType.Bool)] out bool pbMute);
        [PreserveSig] int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
        [PreserveSig] int VolumeStepUp(ref Guid pguidEventContext);
        [PreserveSig] int VolumeStepDown(ref Guid pguidEventContext);
        [PreserveSig] int QueryHardwareSupport(out uint pdwHardwareSupportMask);
        [PreserveSig] int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
    }
    
    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    class MMDeviceEnumeratorComObject { }

    public class PolicyConfigClient {
        [DllImport("ole32.dll")] private static extern int CLSIDFromString([MarshalAs(UnmanagedType.LPWStr)] string lpsz, out Guid pclsid);
        [DllImport("ole32.dll")] private static extern int CoCreateInstance(ref Guid rclsid, IntPtr pUnkOuter, uint dwClsContext, ref Guid riid, out IntPtr ppv);
        [Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPolicyConfig {
            [PreserveSig] int GetMixFormat(string pszDeviceName, IntPtr ppFormat);
            [PreserveSig] int GetDeviceFormat(string pszDeviceName, bool bDefault, IntPtr ppFormat);
            [PreserveSig] int ResetDeviceFormat(string pszDeviceName);
            [PreserveSig] int SetDeviceFormat(string pszDeviceName, IntPtr pEndpointFormat, IntPtr MixFormat);
            [PreserveSig] int GetProcessingPeriod(string pszDeviceName, bool bDefault, IntPtr pmftDefaultPeriod, IntPtr pmftMinimumPeriod);
            [PreserveSig] int SetProcessingPeriod(string pszDeviceName, IntPtr pmftPeriod);
            [PreserveSig] int GetShareMode(string pszDeviceName, IntPtr pMode);
            [PreserveSig] int SetShareMode(string pszDeviceName, IntPtr mode);
            [PreserveSig] int GetPropertyValue(string pszDeviceName, bool bFxStore, ref PropertyKey key, IntPtr pv);
            [PreserveSig] int SetPropertyValue(string pszDeviceName, bool bFxStore, ref PropertyKey key, IntPtr pv);
            [PreserveSig] int SetDefaultEndpoint(string pszDeviceName, int role);
            [PreserveSig] int SetEndpointVisibility(string pszDeviceName, bool bVisible);
        }
        [StructLayout(LayoutKind.Sequential)] private struct PropertyKey { public Guid fmtid; public uint pid; }
        public static void SetDefaultDevice(string deviceId, bool setCommunications) {
            try {
                Guid clsid; Guid iid = typeof(IPolicyConfig).GUID;
                CLSIDFromString("{870af99c-171d-4f9e-af0d-e63df40c2bc9}", out clsid);
                IntPtr pUnknown;
                int hr = CoCreateInstance(ref clsid, IntPtr.Zero, 1, ref iid, out pUnknown);
                if (hr == 0 && pUnknown != IntPtr.Zero) {
                    IPolicyConfig policyConfig = (IPolicyConfig)Marshal.GetObjectForIUnknown(pUnknown);
                    policyConfig.SetDefaultEndpoint(deviceId, 0);
                    policyConfig.SetDefaultEndpoint(deviceId, 1);
                    if (setCommunications) { policyConfig.SetDefaultEndpoint(deviceId, 2); }
                    Marshal.Release(pUnknown);
                }
            } catch (Exception) { }
        }
    }

    public class VolumeToggle {
        public static bool ToggleMute(bool isCapture) {
            IMMDeviceEnumerator deviceEnumerator = null;
            IMMDevice device = null;
            IAudioEndpointVolume endpointVolume = null;
            try {
                deviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
                // 0 = eRender (Speakers), 1 = eCapture (Mic)
                int flow = isCapture ? 1 : 0;
                deviceEnumerator.GetDefaultAudioEndpoint(flow, 0, out device);
                Guid IID_IAudioEndpointVolume = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
                object obj = null;
                device.Activate(ref IID_IAudioEndpointVolume, 23, IntPtr.Zero, out obj);
                endpointVolume = (IAudioEndpointVolume)obj;
                bool isMuted;
                endpointVolume.GetMute(out isMuted);
                Guid ctx = Guid.Empty;
                endpointVolume.SetMute(!isMuted, ref ctx);
                return !isMuted; // returns current NEW mute state
            } finally {
                if (endpointVolume != null) Marshal.ReleaseComObject(endpointVolume);
                if (device != null) Marshal.ReleaseComObject(device);
                if (deviceEnumerator != null) Marshal.ReleaseComObject(deviceEnumerator);
            }
        }
    }
}
'@

function Get-Devices {
    $pnp = Get-PnpDevice -Class "AudioEndpoint" -Status OK -ErrorAction SilentlyContinue
    $results = @()
    foreach ($d in $pnp) {
        $type = $null
        if ($d.InstanceId -like "*`{0.0.0.00000000`}.*") { $type = "Output" }
        elseif ($d.InstanceId -like "*`{0.0.1.00000000`}.*") { $type = "Input" }
        if ($type) {
            $results += [PSCustomObject]@{
                Type = $type
                Name = $d.FriendlyName
                ID   = ($d.InstanceId -replace 'SWD\\MMDEVAPI\\', '')
            }
        }
    }
    return $results
}

$msg = ""
$showInputs = $false
$needsRedraw = $true

# Path to the specific Windows Proximity sound
$proxSound = "C:\Windows\Media\Windows Proximity Notification.wav"
$unmuteSound = "C:\Windows\Media\Windows Unlock.wav"
$muteSound = "C:\Windows\Media\Windows Navigation Start.wav"

while ($true) {
    if ($needsRedraw) {
        Clear-Host
        Write-Host "====================================================" -ForegroundColor DarkCyan
        Write-Host "           SHELL AUDIO DEVICE MANAGER               " -ForegroundColor Cyan
        Write-Host "====================================================" -ForegroundColor DarkCyan
        
        if ($msg) { 
            Write-Host ""
            Write-Host "  >>> $msg" -ForegroundColor Green
            $msg = "" 
        }

        $all = Get-Devices
        $global:activeDevices = @{}
        $idx = 1

        if (-not $showInputs) {
            Write-Host "`n  --- [ OUTPUTS ] (Speakers/Headphones) ---" -ForegroundColor Yellow
            Write-Host "  ------------------------------------------" -ForegroundColor DarkGray
            foreach ($dev in ($all | Where-Object { $_.Type -eq "Output" })) {
                Write-Host "   [" -NoNewline -ForegroundColor Gray
                Write-Host "$idx" -NoNewline -ForegroundColor Green
                Write-Host "] " -NoNewline -ForegroundColor Gray
                Write-Host "$($dev.Name)" -ForegroundColor White
                $global:activeDevices[$idx] = $dev
                $idx++
                if ($idx -gt 9) { break }
            }
        } else {
            Write-Host "`n  --- [ INPUTS ] (Microphones) ---" -ForegroundColor Yellow
            Write-Host "  ------------------------------------------" -ForegroundColor DarkGray
            foreach ($dev in ($all | Where-Object { $_.Type -eq "Input" })) {
                Write-Host "   [" -NoNewline -ForegroundColor Gray
                Write-Host "$idx" -NoNewline -ForegroundColor DarkYellow
                Write-Host "] " -NoNewline -ForegroundColor Gray
                Write-Host "$($dev.Name)" -ForegroundColor White
                $global:activeDevices[$idx] = $dev
                $idx++
                if ($idx -gt 9) { break }
            }
        }

        Write-Host "`n  ==========================================" -ForegroundColor DarkCyan
        Write-Host "   CONTROLS:" -ForegroundColor Gray
        Write-Host "   [1-9] " -NoNewline -ForegroundColor Green
        Write-Host "Select Device" -ForegroundColor White
        Write-Host "   [TAB] " -NoNewline -ForegroundColor Cyan
        Write-Host "Switch Category" -ForegroundColor Gray
        Write-Host "   [M]   " -NoNewline -ForegroundColor Yellow
        Write-Host "Mute Mic" -ForegroundColor Gray
        Write-Host "   [K]   " -NoNewline -ForegroundColor Yellow
        Write-Host "Mute Audio" -ForegroundColor Gray
        Write-Host "   [ESC] " -NoNewline -ForegroundColor Red
        Write-Host "Exit" -ForegroundColor Gray
        Write-Host "  ==========================================" -ForegroundColor Yellow
        Write-Host "   Created by: " -NoNewline -ForegroundColor Gray
        Write-Host "esershnr" -ForegroundColor Cyan
        Write-Host "   GitHub: " -NoNewline -ForegroundColor Gray
        Write-Host "https://github.com/esershnr" -ForegroundColor Blue
        Write-Host "  ==========================================" -ForegroundColor Yellow
        
        $needsRedraw = $false
    }

    if ($Host.UI.RawUI.KeyAvailable) { $null = $Host.UI.RawUI.FlushInputBuffer() }
    $key = [System.Console]::ReadKey($true)
    
    if ($key.Key -eq "Escape") { break }
    if ($key.Key -eq "Tab") { $showInputs = -not $showInputs; $needsRedraw = $true; continue }
    
    # MIC MUTE TOGGLE (M tuşu) - isCapture = true
    if ($key.Key -eq "M") {
        $isMuted = [AudioSwitcher.VolumeToggle]::ToggleMute($true)
        # Audio feedback handled based on new state
        try {
            if ($isMuted) {
                if (Test-Path $muteSound) { (New-Object System.Media.SoundPlayer $muteSound).Play() }
                $msg = "Microphone MUTED"
            } else {
                if (Test-Path $unmuteSound) { (New-Object System.Media.SoundPlayer $unmuteSound).Play() }
                $msg = "Microphone UNMUTED"
            }
        } catch {}
        $needsRedraw = $true
        continue
    }

    # AUDIO MUTE TOGGLE (K tuşu) - isCapture = false
    if ($key.Key -eq "K") {
        $isMuted = [AudioSwitcher.VolumeToggle]::ToggleMute($false)
        # Audio feedback: Only play sound if UNMUTED, because if muted we can't hear it properly! 
        # But actually, users might want to hear the 'mute' sound BEFORE it cuts off, or just visual feedback.
        # System behavior usually plays sound then mutes. Let's try to play sound regardless.
        try {
            if ($isMuted) {
                # Muted state
                $msg = "Audio Output MUTED"
            } else {
                # Unmuted state
                 if (Test-Path $unmuteSound) { (New-Object System.Media.SoundPlayer $unmuteSound).Play() }
                $msg = "Audio Output UNMUTED"
            }
        } catch {}
        $needsRedraw = $true
        continue
    }

    $keyStr = $key.Key.ToString()
    $num = -1
    if ($keyStr -match "D(\d+)") { $num = [int]$matches[1] }
    elseif ($keyStr -match "NumPad(\d+)") { $num = [int]$matches[1] }

    if ($num -gt 0 -and $num -lt 10 -and $global:activeDevices.ContainsKey($num)) {
        $target = $global:activeDevices[$num]
        [AudioSwitcher.PolicyConfigClient]::SetDefaultDevice($target.ID, $true)
        try {
            if (Test-Path $proxSound) { (New-Object System.Media.SoundPlayer $proxSound).Play() }
        } catch {}
        $msg = "Switched to: $($target.Name)"
        $needsRedraw = $true
    }
}

Clear-Host
Write-Host "Exited. Made by esershnr." -ForegroundColor Gray
