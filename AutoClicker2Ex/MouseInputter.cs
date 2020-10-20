using System;
using System.Runtime.InteropServices;
using System.Threading;
using WindowsInput;
using WindowsInput.Native;

namespace Auto_Clicker
{
    public class MouseInputter
    {
        //Import unmanaged functions from DLL library
        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SendInput(int nInputs, ref MOUSEINPUT pInputs, int cbSize);

        /// <summary>
        /// Structure for SendInput function holding relevant mouse coordinates and information
        /// </summary>
        private struct MOUSEINPUT
        {
            public uint type;
            public MOUSEINPUTDATA mi;
        };

        /// <summary>
        /// Structure for SendInput function holding coordinates of the click and other information
        /// </summary>
        private struct MOUSEINPUTDATA
        {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        };

        //Constants for use in SendInput and mouse_event
        public const int INPUT_MOUSE = 0x0000;

        public const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        public const int MOUSEEVENTF_LEFTUP = 0x0004;
        public const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        public const int MOUSEEVENTF_RIGHTUP = 0x0010;
        public const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        public const int MOUSEEVENTF_MIDDLEUP = 0x0040;

        #region Mouse_Event Methods

        /// <summary>
        /// Click the left mouse button at the current cursor position using
        /// the imported mouse_event function
        /// </summary>
        private void ClickLeftMouseButtonMouseEvent()
        {
            //Send a left click down followed by a left click up to simulate a 
            //full left click
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }

        /// <summary>
        /// Click the right mouse button at the current cursor position using
        /// the imported mouse_event function
        /// </summary>
        private void ClickRightMouseButtonMouseEvent()
        {
            //Send a left click down followed by a right click up to simulate a 
            //full right click
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
        }

        #endregion

        #region SendInput Methods

        /// <summary>
        /// Click the left mouse button at the current cursor position using
        /// the imported SendInput function
        /// </summary>
        public static void ClickLeftMouseButtonSendInput()
        {
            //Initialise INPUT object with corresponding values for a left click
            MOUSEINPUT mouseinput = new MOUSEINPUT();
            mouseinput.type = INPUT_MOUSE;
            mouseinput.mi.dx = 0;
            mouseinput.mi.dy = 0;
            mouseinput.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
            mouseinput.mi.dwExtraInfo = IntPtr.Zero;
            mouseinput.mi.mouseData = 0;
            mouseinput.mi.time = 0;

            //Send a left click down followed by a left click up to simulate a 
            //full left click
            SendInput(1, ref mouseinput, Marshal.SizeOf(mouseinput));
            Thread.Sleep(10); // Need for Windows to recognize a click
            mouseinput.mi.dwFlags = MOUSEEVENTF_LEFTUP;
            SendInput(1, ref mouseinput, Marshal.SizeOf(mouseinput));
        }

        /// <summary>
        /// Click the left mouse button at the current cursor position using
        /// the imported SendInput function
        /// </summary>
        public static void ClickRightMouseButtonSendInput()
        {
            //Initialise INPUT object with corresponding values for a right click
            MOUSEINPUT mouseinput = new MOUSEINPUT();
            mouseinput.type = INPUT_MOUSE;
            mouseinput.mi.dx = 0;
            mouseinput.mi.dy = 0;
            mouseinput.mi.dwFlags = MOUSEEVENTF_RIGHTDOWN;
            mouseinput.mi.dwExtraInfo = IntPtr.Zero;
            mouseinput.mi.mouseData = 0;
            mouseinput.mi.time = 0;

            //Send a right click down followed by a right click up to simulate a 
            //full right click
            SendInput(1, ref mouseinput, Marshal.SizeOf(mouseinput));
            Thread.Sleep(10); // Need for Windows to recognize a click
            mouseinput.mi.dwFlags = MOUSEEVENTF_RIGHTUP;
            SendInput(1, ref mouseinput, Marshal.SizeOf(mouseinput));
        }

        public static void ClickMiddleMouseButtonSendInput()
        {
            InputSimulator s = new InputSimulator();

            //Initialise INPUT object with corresponding values for a right click
            MOUSEINPUT mouseinput = new MOUSEINPUT();
            mouseinput.type = INPUT_MOUSE;
            mouseinput.mi.dx = 0;
            mouseinput.mi.dy = 0;
            mouseinput.mi.dwFlags = MOUSEEVENTF_MIDDLEDOWN;
            mouseinput.mi.dwExtraInfo = IntPtr.Zero;
            mouseinput.mi.mouseData = 0;
            mouseinput.mi.time = 0;

            //Send a right click down followed by a right click up to simulate a 
            //full right click
            SendInput(1, ref mouseinput, Marshal.SizeOf(mouseinput));
            Thread.Sleep(10); // Need for Windows to recognize a click
            mouseinput.mi.dwFlags = MOUSEEVENTF_MIDDLEUP;
            SendInput(1, ref mouseinput, Marshal.SizeOf(mouseinput));
        }

        public void ClickLetterZ()
        {
            InputSimulator s = new InputSimulator();
            s.Keyboard.KeyDown(VirtualKeyCode.VK_Z);
            s.Keyboard.Sleep(10);
            s.Keyboard.KeyUp(VirtualKeyCode.VK_Z);
        }

        public void ClickEnterKey()
        {
            InputSimulator s = new InputSimulator();
            s.Keyboard.KeyDown(VirtualKeyCode.RETURN);
            s.Keyboard.Sleep(10);
            s.Keyboard.KeyUp(VirtualKeyCode.RETURN);
        }

        #endregion
    }
}