using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CoreAudio;

namespace ClassicVolumeMixer.Helpers
{
    public static class AudioDeviceHelper
    {
        public static IEnumerable<MMDevice> GetAudioDevices(DataFlow dataFlow)
        {
            return new MMDeviceEnumerator(new Guid()).EnumerateAudioEndPoints(dataFlow, DeviceState.Active);
        }

        public static bool IsAudioDeviceAvailable()
        {
            return GetDefaultAudioDevice() != null;
        }

        public static MMDevice GetDefaultAudioDevice()
        {
            try
            {
                return new MMDeviceEnumerator(new Guid()).GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            }
            catch
            {
                return null;
            }
        }

        public static void SetDefaultAudioDevice(MMDevice device)
        {
            new CoreAudio.CPolicyConfigVistaClient().SetDefaultDevice(device.ID);
        }
        public static int GetVolumeLevel()
        {
            MMDevice defaultAudioDevice = GetDefaultAudioDevice();
            return (int)(defaultAudioDevice.AudioEndpointVolume.MasterVolumeLevelScalar * 100);
        }
        public static bool IsMuted()
        {
            MMDevice defaultAudioDevice = GetDefaultAudioDevice();
            return defaultAudioDevice.AudioEndpointVolume.Mute;
        }
    }
}