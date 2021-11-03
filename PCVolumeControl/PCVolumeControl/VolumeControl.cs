using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
//custom Namespaces
using VolumeControl.Library.Constants;
using VolumeControl.Library.Structs;
using VolumeControl.Library.Win32;

namespace VolumeControl.Library.Constants
{
    /// <summary>
    /// Class to hold all the constants needed for controlling the system sound
    /// </summary>
    public static class VolumeConstants
    {
        public const int MMSYSERR_NOERROR = 0;
        public const int MAXPNAMELEN = 32;
        public const int MIXER_LONG_NAME_CHARS = 64;
        public const int MIXER_SHORT_NAME_CHARS = 16;
        public const int MIXER_GETLINEINFOF_COMPONENTTYPE = 0x3;
        public const int MIXER_GETCONTROLDETAILSF_VALUE = 0x0;
        public const int MIXER_GETLINECONTROLSF_ONEBYTYPE = 0x2;
        public const int MIXER_SETCONTROLDETAILSF_VALUE = 0x0;
        public const int MIXERLINE_COMPONENTTYPE_DST_FIRST = 0x0;
        public const int MIXERLINE_COMPONENTTYPE_SRC_FIRST = 0x1000;        
        public const int MIXERCONTROL_CT_CLASS_FADER = 0x50000000;
        public const int MIXERCONTROL_CT_UNITS_UNSIGNED = 0x30000;
        public const int MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER | MIXERCONTROL_CT_UNITS_UNSIGNED);
        public const int MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1);
        public const int MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4);
        public const int MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3);
        public const int MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2);
    }
}
namespace VolumeControl.Library.Structs
{
    /// <summary>
    /// Class file for holding all the custom sructures we need
    /// for controlling the system sound
    /// </summary>
    public static class VolumeStructs
    {
        /// <summary>
        /// struct for holding data for the mixer caps
        /// </summary>
        public struct MixerCaps
        {
            public int wMid;
            public int wPid;
            public int vDriverVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = VolumeConstants.MAXPNAMELEN)] public string szPname;
            public int fdwSupport;
            public int cDestinations;
        }

        /// <summary>
        /// struct to hold data for the mixer control
        /// </summary>
        public struct Mixer
        {
            public int cbStruct;
            public int dwControlID;
            public int dwControlType;
            public int fdwControl;
            public int cMultipleItems;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = VolumeConstants.MIXER_SHORT_NAME_CHARS)] public string szShortName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = VolumeConstants.MIXER_LONG_NAME_CHARS)] public string szName;
            public int lMinimum;
            public int lMaximum;
            [MarshalAs(UnmanagedType.U4, SizeConst = 10)]  public int reserved;
        }

        /// <summary>
        /// struct for holding data about the details of the mixer control
        /// </summary>
        public struct MixerDetails
        {
            public int cbStruct;
            public int dwControlID;
            public int cChannels;
            public int item;
            public int cbDetails;
            public IntPtr paDetails;
        }

        /// <summary>
        /// struct to hold data for an unsigned mixer control details
        /// </summary>
        public struct UnsignedMixerDetails
        {
            public int dwValue;
        }

        /// <summary>
        /// struct to hold data for the mixer line
        /// </summary>
        public struct MixerLine
        {
            public int cbStruct;
            public int dwDestination;
            public int dwSource;
            public int dwLineID;
            public int fdwLine;
            public int dwUser;
            public int dwComponentType;
            public int cChannels;
            public int cConnections;
            public int cControls;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = VolumeConstants.MIXER_SHORT_NAME_CHARS)] public string szShortName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = VolumeConstants.MIXER_LONG_NAME_CHARS)] public string szName;
            public int dwType;
            public int dwDeviceID;
            public int wMid;
            public int wPid;
            public int vDriverVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = VolumeConstants.MAXPNAMELEN)] public string szPname;
        }

        /// <summary>
        /// struct for holding data for the mixer line controls
        /// </summary>
        public struct LineControls
        {
            public int cbStruct;
            public int dwLineID;
            public int dwControl;
            public int cControls;
            public int cbmxctrl;
            public IntPtr pamxctrl;
        }
    }
}
namespace VolumeControl.Library.Win32
{
    /// <summary>
    /// Class that contains all the Win32 API calls
    /// we need for controlling the system sound
    /// </summary>
    public static class PCWin32
    {
        [DllImport("winmm.dll", CharSet=CharSet.Ansi)] 
        public static extern int mixerClose (int hmx);

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerGetControlDetailsA(int hmxobj, ref VolumeStructs.MixerDetails pmxcd, int fdwDetails);

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerGetDevCapsA(int uMxId, VolumeStructs.MixerCaps pmxcaps, int cbmxcaps);

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerGetID(int hmxobj, int pumxID, int fdwId);

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerGetLineControlsA(int hmxobj, ref VolumeStructs.LineControls pmxlc, int fdwControls);

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerGetLineInfoA(int hmxobj, ref VolumeStructs.MixerLine pmxl, int fdwInfo);

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerGetNumDevs();

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerMessage(int hmx, int uMsg, int dwParam1, int dwParam2);

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerOpen(out int phmx, int uMxId, int dwCallback, int dwInstance, int fdwOpen);

        [DllImport("winmm.dll", CharSet=CharSet.Ansi)]
        public static extern int mixerSetControlDetails(int hmxobj, ref VolumeStructs.MixerDetails pmxcd, int fdwDetails);
    }
}

namespace PC_VolumeControl
{
    public class VolumeControl
    {
        /// <summary>
        /// method to retrieve a mixer control
        /// </summary>
        /// <param name="i"></param>
        /// <param name="type"></param>
        /// <param name="ctrlType"></param>
        /// <param name="mixerControl"></param>
        /// <param name="currVolume"></param>
        /// <returns></returns>
        private static bool GetMixer(int i, int type, int ctrlType, out VolumeStructs.Mixer mixerControl, out int currVolume)
        {
            //create our method level variables
            int details;
            bool bReturn;
            currVolume = -1;

            //create our struct objects
            VolumeStructs.LineControls lineControls = new VolumeStructs.LineControls();
            VolumeStructs.MixerLine line = new VolumeStructs.MixerLine();
            VolumeStructs.MixerDetails mcDetails = new VolumeStructs.MixerDetails();
            VolumeStructs.UnsignedMixerDetails detailsUnsigned = new VolumeStructs.UnsignedMixerDetails();

            //create a new mixer control
            mixerControl = new VolumeStructs.Mixer();
           
            //set the properties of out mixerline object
            line.cbStruct = Marshal.SizeOf(line);
            line.dwComponentType = type;
            //get the line info and assign it to our details variable
            details = PCWin32.mixerGetLineInfoA(i, ref line, VolumeConstants.MIXER_GETLINEINFOF_COMPONENTTYPE);

            //make sure we didnt receive any errors
            if (VolumeConstants.MMSYSERR_NOERROR == details)
            {
                int mcSize = 152;
                //get the size of the unmanaged type
                int control = Marshal.SizeOf(typeof(VolumeStructs.Mixer));
                //allocate a block of memory
                lineControls.pamxctrl = Marshal.AllocCoTaskMem(mcSize);
                //get the size of the line controls
                lineControls.cbStruct = Marshal.SizeOf(lineControls);

                //set properties for our mixer control
                lineControls.dwLineID = line.dwLineID;
                lineControls.dwControl = ctrlType;
                lineControls.cControls = 1;
                lineControls.cbmxctrl = mcSize;

                // Allocate a buffer for the control
                mixerControl.cbStruct = mcSize;

                // Get the control
                details = PCWin32.mixerGetLineControlsA(i, ref lineControls, VolumeConstants.MIXER_GETLINECONTROLSF_ONEBYTYPE);

                //once again check to see if we received any errors
                if (VolumeConstants.MMSYSERR_NOERROR == details)
                {
                    bReturn = true;
                    //Copy the control into the destination structure
                    mixerControl = (VolumeStructs.Mixer)Marshal.PtrToStructure(lineControls.pamxctrl, typeof(VolumeStructs.Mixer));
                }
                else
                {
                    bReturn = false;
                }

                int mcDetailsSize = Marshal.SizeOf(typeof(VolumeStructs.MixerDetails));
                int mcDetailsUnsigned = Marshal.SizeOf(typeof(VolumeStructs.UnsignedMixerDetails));
                mcDetails.cbStruct = mcDetailsSize;
                mcDetails.dwControlID = mixerControl.dwControlID;
                mcDetails.paDetails = Marshal.AllocCoTaskMem(mcDetailsUnsigned);
                mcDetails.cChannels = 1;
                mcDetails.item = 0;
                mcDetails.cbDetails = mcDetailsUnsigned;
                details = PCWin32.mixerGetControlDetailsA(i, ref mcDetails, VolumeConstants.MIXER_GETCONTROLDETAILSF_VALUE);
                detailsUnsigned = (VolumeStructs.UnsignedMixerDetails)Marshal.PtrToStructure(mcDetails.paDetails, typeof(VolumeStructs.UnsignedMixerDetails));
                currVolume = detailsUnsigned.dwValue;
                return bReturn;
            }

            bReturn = false;
            return bReturn;
        }


        /// <summary>
        /// method for setting the value for a volume control
        /// </summary>
        /// <param name="i"></param>
        /// <param name="mixerControl"></param>
        /// <param name="volumeLevel"></param>
        /// <returns>true/false</returns>
        private static bool SetMixer(int i, VolumeStructs.Mixer mixerControl, int volumeLevel)
        {
            //method level variables
            bool bReturn;
            int details;

            //create our struct object for controlling the system sound
            VolumeStructs.MixerDetails mixerDetails = new VolumeStructs.MixerDetails();
            VolumeStructs.UnsignedMixerDetails volume = new VolumeStructs.UnsignedMixerDetails();

            //set out mixer control properties
            mixerDetails.item = 0;
            //set the id of the mixer control
            mixerDetails.dwControlID = mixerControl.dwControlID;
            //return the size of the mixer details struct
            mixerDetails.cbStruct = Marshal.SizeOf(mixerDetails);
            //return the volume
            mixerDetails.cbDetails = Marshal.SizeOf(volume);

            //Allocate a buffer for the mixer control value buffer
            mixerDetails.cChannels = 1;
            volume.dwValue = volumeLevel;

            //Copy the data into the mixer control value buffer
            mixerDetails.paDetails = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(VolumeStructs.UnsignedMixerDetails)));
            Marshal.StructureToPtr(volume, mixerDetails.paDetails, false);

            //Set the control value
            details = PCWin32.mixerSetControlDetails(i, ref mixerDetails, VolumeConstants.MIXER_SETCONTROLDETAILSF_VALUE);

            //Check to see if any errors were returned
            if (VolumeConstants.MMSYSERR_NOERROR == details)
            {
                bReturn = true;
            }
            else
            {
                bReturn = false;
            }
            return bReturn;

        }


        /// <summary>
        /// method for retrieving the current volume from the system
        /// </summary>
        /// <returns>int value</returns>
        public static int GetVolume()
        {
            //method level variables
            int currVolume;
            int mixerControl;

            //create a new volume control
            VolumeStructs.Mixer mixer = new VolumeStructs.Mixer();
            
            //open the mixer
            PCWin32.mixerOpen(out mixerControl, 0, 0, 0, 0);

            //set the type to volume control type
            int type = VolumeConstants.MIXERCONTROL_CONTROLTYPE_VOLUME;

            //get the mixer control and get the current volume level
            GetMixer(mixerControl, VolumeConstants.MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, type, out mixer, out currVolume);

            //close the mixer control since we are now done with it
            PCWin32.mixerClose(mixerControl);

            //return the current volume to the calling method
            return currVolume;
        }


        /// <summary>
        /// method for setting the volume to a specific level
        /// </summary>
        /// <param name="volumeLevel">volume level we wish to set volume to</param>
        public static void SetVolume(int volumeLevel)
        {
            try
            {
                //method level variables
                int currVolume;
                int mixerControl;

                //create a new volume control
                VolumeStructs.Mixer volumeControl = new VolumeStructs.Mixer();

                //open the mixer control
                PCWin32.mixerOpen(out mixerControl, 0, 0, 0, 0);

                //set the type to volume control type
                int controlType = VolumeConstants.MIXERCONTROL_CONTROLTYPE_VOLUME;

                //get the current mixer control and get the current volume
                GetMixer(mixerControl, VolumeConstants.MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, controlType, out volumeControl, out currVolume);

                //now check the volume level. If the volume level
                //is greater than the max then set the volume to
                //the max level. If it's less than the minimum level
                //then set it to the minimun level
                if (volumeLevel > volumeControl.lMaximum)
                {
                    volumeLevel = volumeControl.lMaximum;
                }
                else if (volumeLevel < volumeControl.lMinimum)
                {
                    volumeLevel = volumeControl.lMinimum;
                }

                //set the volume
                SetMixer(mixerControl, volumeControl, volumeLevel);

                //now re-get the mixer control
                GetMixer(mixerControl, VolumeConstants.MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, controlType, out volumeControl, out currVolume);

                //make sure the volume level is equal to the current volume
                if (volumeLevel != currVolume)
                {
                    throw new Exception("Cannot Set Volume");
                }

                //close the mixer control as we are finished with it
                PCWin32.mixerClose(mixerControl);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }
    }
}
