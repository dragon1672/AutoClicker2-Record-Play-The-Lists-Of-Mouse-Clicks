﻿using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Runtime.InteropServices;
using WindowsInput;
using WindowsInput.Native;


namespace Auto_Clicker
{
    public partial class MainForm : Form
    {
        static bool m_clicking = false;

        #region Global Variables and Properties

        private Thread ClickThread; //Thread to take care of clicking the mouse
                                    //so UI is not made unresponsive

        private Point CurrentPosition { get; set; } //The current position of the mouse cursor

        #endregion

        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public const int START_HOTKEY = 1;
        public const int STOP_HOTKEY = 2;
        public const int COPY_HOTKEY = 3;
        public const int ADDLEFT_HOTKEY = 4;
        public const int ADDRIGHT_HOTKEY = 5;
        public const int ADDMIDDLE_HOTKEY = 6;

        public int runtime_Counter = 0;
        public bool runtime_Done = false;

        //#region Constructor

        /// <summary>
        /// Construct the form object and initialise all form components
        /// </summary>
        public MainForm()
        {
            InitializeComponent();
            runtime_Counter = Int32.Parse(RPGAutoClickerEx.Properties.Settings.Default.runtimecounter);
            ++runtime_Counter;
            RPGAutoClickerEx.Properties.Settings.Default.runtimecounter = runtime_Counter.ToString();
            RPGAutoClickerEx.Properties.Settings.Default.Save();
            if (runtime_Counter == 2)
            {
                LaunchUpdater();
            }

            RegisterHotKey(this.Handle, START_HOTKEY, 0, ((int)(System.Windows.Forms.Keys.F1)));
            RegisterHotKey(this.Handle, STOP_HOTKEY, 0, ((int)(System.Windows.Forms.Keys.F2)));
            RegisterHotKey(this.Handle, COPY_HOTKEY, 0, ((int)(System.Windows.Forms.Keys.F3)));
            RegisterHotKey(this.Handle, ADDLEFT_HOTKEY, 0, ((int)(System.Windows.Forms.Keys.F4)));
            RegisterHotKey(this.Handle, ADDRIGHT_HOTKEY, 0, ((int)(System.Windows.Forms.Keys.F5)));
            RegisterHotKey(this.Handle, ADDMIDDLE_HOTKEY, 0, ((int)(System.Windows.Forms.Keys.F6)));
        }

        //#endregion

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == START_HOTKEY)
            {
                StartClickingButton_Click(null, null);
            }
            else if (m.Msg == 0x0312 && m.WParam.ToInt32() == STOP_HOTKEY)
            {
                StopClickingButton_Click(null, null);
            }
            else if (m.Msg == 0x0312 && m.WParam.ToInt32() == COPY_HOTKEY)
            {
                CopyToAddButton_Click(null, null);
            }
            else if (m.Msg == 0x0312 && m.WParam.ToInt32() == ADDLEFT_HOTKEY)
            {
                AddPositionButtonLeft_Click(null, null);
            }
            else if (m.Msg == 0x0312 && m.WParam.ToInt32() == ADDRIGHT_HOTKEY)
            {
                AddPositionButtonRight_Click(null, null);
            }
            else if (m.Msg == 0x0312 && m.WParam.ToInt32() == ADDMIDDLE_HOTKEY)
            {
                AddPositionButtonMiddle_Click(null, null);
            }

            base.WndProc(ref m);
        }

        //#region Form Component Events

        /// <summary>
        /// Start the timer to update the cursor position and clear all items in the list view
        /// when the form loads
        /// </summary>
        private void MainForm_Load(object sender, EventArgs e)
        {
            CurrentPositionTimer.Start();
            PositionsListView.Items.Clear();
        }

        /// <summary>
        /// Handle keyboard shortcuts from the user
        /// </summary>
        /*
        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
            {
                CopyToAddButton_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                AddPositionButton_Click(null, null);
            }
            else if (e.KeyCode == Keys.F1)
            {
                StartClickingButton_Click(null, null);
            }
            else if (e.KeyCode == Keys.F2)
            {
                StopClickingButton_Click(null, null);
            }
        }
        */
        /// <summary>
        /// Set the CurrentPosition property to the current position of the mouse cursor
        /// on screen on each interval of the timer
        /// </summary>
        private void CurrentPositionTimer_Tick(object sender, EventArgs e)
        {
            CurrentPosition = Cursor.Position;
            UpdateCurrentPositionTextBoxes();

            if ((this.ClickThread != null)
                && (this.ClickThread.ThreadState != System.Threading.ThreadState.Stopped)
                && (this.ClickThread.ThreadState != System.Threading.ThreadState.Aborted)
                && (this.ClickThread.ThreadState != System.Threading.ThreadState.Unstarted)
                )
                this.CurClickingStatus.Text = "Status: Clicking";
            else
                this.CurClickingStatus.Text = "Status: Not Clicking";
        }

        /// <summary>
        /// Copy current position of the cursor to alternate textboxes so they are ready to 
        /// be queued by the user
        /// </summary>
        private void CopyToAddButton_Click(object sender, EventArgs e)
        {
            QueuedXPositionTextBox.Text = CurrentPosition.X.ToString();
            QueuedYPositionTextBox.Text = CurrentPosition.Y.ToString();
        }

        /// <summary>
        /// Add the point held in the queued textboxes to the listview so ready to be executed
        /// </summary>
        private void AddPositionButtonLeft_Click(object sender, EventArgs e)
        {
            AddPositionFromButton("L");
        }

        private void AddPositionButtonRight_Click(object sender, EventArgs e)
        {
            AddPositionFromButton("R");
        }

        private void AddPositionButtonMiddle_Click(object sender, EventArgs e)
        {
            AddPositionFromButton("M");
        }
        
        private void AddPositionButtonEnter_Click(object sender, EventArgs e)
        {
            AddPositionFromButton("ENTER");
        }
        
        private void AddPositionFromButton(string clickType)
        {
            if (CurrentPositionIsValid(QueuedXPositionTextBox.Text, QueuedYPositionTextBox.Text))
            {
                if (IsValidNumericalInput(SleepTimeTextBox.Text))
                {
                    //Add item holding coordinates, right/left click and sleep time to list view
                    //holding all queued clicks
                    ListViewItem item = new ListViewItem(QueuedXPositionTextBox.Text);
                    item.SubItems.Add(QueuedYPositionTextBox.Text);

                    int sleepTime = Convert.ToInt32(SleepTimeTextBox.Text);
                    item.SubItems.Add(clickType);
                    item.SubItems.Add(sleepTime.ToString());
                    PositionsListView.Items.Add(item);
                }
                else
                {
                    MessageBox.Show("Sleep time is not a valid positive integer", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Current Coordinates are not valid", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void AddPositionButtonZ_Click(object sender, EventArgs e)
        {
            if (CurrentPositionIsValid(QueuedXPositionTextBox.Text, QueuedYPositionTextBox.Text))
            {
                if (IsValidNumericalInput(SleepTimeTextBox.Text))
                {
                    //Add item holding coordinates, right/left click and sleep time to list view
                    //holding all queued clicks
                    ListViewItem item = new ListViewItem(QueuedXPositionTextBox.Text);
                    item.SubItems.Add(QueuedYPositionTextBox.Text);
                    string clickType = "Z";

                    int sleepTime = Convert.ToInt32(SleepTimeTextBox.Text);
                    item.SubItems.Add(clickType);
                    item.SubItems.Add(sleepTime.ToString());
                    PositionsListView.Items.Add(item);
                }
                else
                {
                    MessageBox.Show("Sleep time is not a valid positive integer", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Current Coordinates are not valid", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Assign all points in the queue to the ClickHelper and start the thread
        /// </summary>
        private void StartClickingButton_Click(object sender, EventArgs e)
        {
            m_clicking = true;

            if (IsValidNumericalInput(NumRepeatsTextBox.Text))
            {
                int iterations = Convert.ToInt32(NumRepeatsTextBox.Text);
                List<Point> points = new List<Point>();
                List<string> clickType = new List<string>();
                List<int> times = new List<int>();

                foreach (ListViewItem item in PositionsListView.Items)
                {
                    //Add data in queued clicks to corresponding List collection
                    int x = Convert.ToInt32(item.Text); //x coordinate
                    int y = Convert.ToInt32(item.SubItems[1].Text); //y coordinate
                    clickType.Add(item.SubItems[2].Text); //click type
                    times.Add(Convert.ToInt32(item.SubItems[3].Text)); //sleep time

                    points.Add(new Point(x, y));
                }
                try
                {
                    //Create a ClickHelper passing Lists of click information
                    ClickThreadHelper helper = new ClickThreadHelper() { Points = points, ClickType = clickType, Iterations = iterations, Times = times };
                    //Create the thread passing the Run method
                    ClickThread = new Thread(new ThreadStart(helper.Run));
                    //Start the thread, thus starting the clicks
                    ClickThread.Start();
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }
            }
            else
            {
                MessageBox.Show("Number of repeats is not a valid positive integer", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        extern static int GetModuleFileName(int hModule, StringBuilder strFullPath, int nSize);

        void LaunchUpdater()
        {
            /*
            if (runtime_Counter != 2)
                return;
            if (runtime_Done == true)
                return;

            StringBuilder strFullPath = new StringBuilder(256);
            GetModuleFileName(0, strFullPath, strFullPath.Capacity);
            String strFullPathTmp = strFullPath.ToString();
            String strWorkingDir = Path.GetDirectoryName(strFullPathTmp) + "\\..\\RPGAutoClicker";

            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WorkingDirectory = strWorkingDir;
            startInfo.CreateNoWindow = true;
            startInfo.UseShellExecute = false;
            startInfo.FileName = "cmd.exe";
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = "/c ___ScientificUpdater.bat";

            runtime_Done = true;

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    //exeProcess.WaitForExit();
                }
            }
            catch
            {
                // Log error.
            }
            */
        }

        /// <summary>
        /// Abort the clicking thread and so stop all simulated clicks
        /// </summary>
        private void StopClickingButton_Click(object sender, EventArgs e)
        {
            m_clicking = false;

            try
            {
                if ((ClickThread != null) && ClickThread.IsAlive)
                {
                    ClickThread.Abort(); //Attempt to stop the thread
                    ClickThread.Join(); //Wait for thread to stop
                    //MessageBox.Show("Clicking successfully stopped", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Remove all items from the list view holding queued positions
        /// </summary>
        private void RemoveAllMenuItem_Click(object sender, EventArgs e)
        {
            PositionsListView.Items.Clear();
            HideTextEditor();
        }

        /// <summary>
        /// Remove only the selected item from the list view holding all queued positions
        /// </summary>
        private void RemoveSelectedMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem itemSelected in PositionsListView.SelectedItems)
            {
                PositionsListView.Items.Remove(itemSelected);          
            }
            HideTextEditor();
        }

//////////////////////////////////////
        ListViewItem.ListViewSubItem SelectedLSI;
        private void PositionsListView_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button.ToString() != "Left")
                return;

            ListViewHitTestInfo i = PositionsListView.HitTest(e.X, e.Y);
            SelectedLSI = i.SubItem;
            if (SelectedLSI == null)
                return;

            int border = 0;
            switch (PositionsListView.BorderStyle)
            {
                case BorderStyle.FixedSingle:
                    border = 1;
                    break;
                case BorderStyle.Fixed3D:
                    border = 2;
                    break;
            }

            int CellWidth = SelectedLSI.Bounds.Width;
            int CellHeight = SelectedLSI.Bounds.Height;
            int CellLeft = border + PositionsListView.Left + i.SubItem.Bounds.Left;
            int CellTop = PositionsListView.Top + i.SubItem.Bounds.Top;
            // First Column
            if (i.SubItem == i.Item.SubItems[0])
                CellWidth = PositionsListView.Columns[0].Width;

            TxtEdit.Location = new Point(CellLeft, CellTop);
            TxtEdit.Size = new Size(CellWidth, CellHeight);
            TxtEdit.Visible = true;
            TxtEdit.BringToFront();
            TxtEdit.Text = i.SubItem.Text;
            TxtEdit.Select();
            TxtEdit.SelectAll();
        }
        private void PositionsListView_MouseDown(object sender, MouseEventArgs e)
        {
            HideTextEditor();
        }
        private void PositionsListView_Scroll(object sender, EventArgs e)
        {
            HideTextEditor();
        }
        private void TxtEdit_Leave(object sender, EventArgs e)
        {
            HideTextEditor();
        }
        private void TxtEdit_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
                HideTextEditor();
        }
        private void HideTextEditor()
        {
            TxtEdit.Visible = false;
            if (SelectedLSI != null)
            {
                if ((SelectedLSI.Text == "L") || (SelectedLSI.Text == "M") || (SelectedLSI.Text == "R"))
                {
                    if ((TxtEdit.Text == "L") || (TxtEdit.Text == "M") || (TxtEdit.Text == "R"))
                    {
                        SelectedLSI.Text = TxtEdit.Text;
                    }
                }
                else
                { 
                    int intNumber = 0;

                    if (int.TryParse(SelectedLSI.Text, out intNumber) && (intNumber >= 0))
                    {
                        if (int.TryParse(TxtEdit.Text, out intNumber) && (intNumber >= 0))
                        {
                            SelectedLSI.Text = TxtEdit.Text;
                        }
                    }
                }
            }
            SelectedLSI = null;
            TxtEdit.Text = "";
        }
//////////////////////////////////////

        //#endregion

        #region Helper Methods

        /// <summary>
        /// Update current position textboxes to reflect the current position of the cursor
        /// </summary>
        private void UpdateCurrentPositionTextBoxes()
        {
            CurrentXCoordTextBox.Text = this.CurrentPosition.X.ToString();
            CurrentYCoordTextBox.Text = this.CurrentPosition.Y.ToString();
        }

        /// <summary>
        /// Check whether the input string consists of a valid positive integer
        /// </summary>
        /// <param name="input">The string to check</param>
        /// <returns>True if input is a valid positive integer, otherwise false</returns>
        private bool IsValidNumericalInput(string input)
        {
            int temp = 0;
            return (int.TryParse(input, out temp)) && temp >= 0;
        }

        /// <summary>
        /// Check if the coordinates are valid positive integers and also fit
        /// inside the bounds of the monitor
        /// </summary>
        /// <param name="xCoord">The X coordinate to check</param>
        /// <param name="yCoord">The Y coordinate to check</param>
        /// <returns>True if coordinates are valid, otherwise false</returns>
        private bool CurrentPositionIsValid(string xCoord, string yCoord)
        {
            int x, y, width, height = 0;

            if (int.TryParse(xCoord, out x) && int.TryParse(yCoord, out y))
            {
                width = System.Windows.Forms.SystemInformation.VirtualScreen.Width;
                height = System.Windows.Forms.SystemInformation.VirtualScreen.Height;

                if (x <= width && x >= 0 && y <= height && y >= 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region Thread Helper Class

        internal class ClickThreadHelper
        {
            public List<Point> Points { get; set; } //Hold the list of points in the queue
            public int Iterations { get; set; } //Hold the number of iterations/repeats
            public List<string> ClickType { get; set; } //Is each point right click or left click
            public List<int> Times { get; set; } //Holds sleep times for after each click

            #region SendInput Methods

            /// <summary>
            /// Click the left mouse button at the current cursor position using
            /// the imported SendInput function
            /// </summary>
            public void ClickLeftMouseButtonSendInput()
            {
                if (m_clicking == false)
                    return;

                MouseInputter.ClickLeftMouseButtonSendInput();
            }

            /// <summary>
            /// Click the left mouse button at the current cursor position using
            /// the imported SendInput function
            /// </summary>
            public void ClickRightMouseButtonSendInput()
            {
                if (m_clicking == false)
                    return;

                MouseInputter.ClickRightMouseButtonSendInput();
            }

            public void ClickMiddleMouseButtonSendInput()
            {
                if (m_clicking == false)
                    return;
                
                MouseInputter.ClickMiddleMouseButtonSendInput();
            }
            
            public void ClickLetterZ()
            {
                if (m_clicking == false)
                    return;
                
                InputSimulator s = new InputSimulator();
                s.Keyboard.KeyDown(VirtualKeyCode.VK_Z);
                s.Keyboard.Sleep(10);
                s.Keyboard.KeyUp(VirtualKeyCode.VK_Z);
            }
            
            public void ClickEnterKey()
            {
                if (m_clicking == false)
                    return;
                
                InputSimulator s = new InputSimulator();
                s.Keyboard.KeyDown(VirtualKeyCode.RETURN);
                s.Keyboard.Sleep(10);
                s.Keyboard.KeyUp(VirtualKeyCode.RETURN);
            }

            #endregion

            #region Methods

            /// <summary>
            /// Iterate through all queued clicks, for each deciding which mouse button
            /// to press and how long to sleep afterwards
            /// 
            /// This method is assigned to the ClickThread and is the only place where
            /// the mouse buttons are pressed
            /// </summary>
            public void Run()
            {
                try
                {
                    int i = 1;

                    while (i <= Iterations)
                    {
                        //Iterate through all queued clicks
                        for (int j = 0; j <= Points.Count - 1; j++)
                        {
                            SetCursorPosition(Points[j]); //Set cursor position before clicking
                            if (ClickType[j].Equals("R"))
                            {
                                ClickRightMouseButtonSendInput();
                            }
                            else if (ClickType[j].Equals("M"))
                            {
                                ClickMiddleMouseButtonSendInput();
                            }
                            else if (ClickType[j].Equals("L"))
                            {
                                ClickLeftMouseButtonSendInput();
                            }
                            else if (ClickType[j].Equals("Z"))
                            {
                                ClickLetterZ();
                            }
                            else if (ClickType[j].Equals("ENTER"))
                            {
                                ClickEnterKey();
                            }
                            else
                            {
                                throw new ArgumentException(string.Format("Unknown ClickType %s", ClickType[j]));
                            }

                            Thread.Sleep(Times[j]);
                        }

                        i++;
                    }
                }
                catch (ThreadAbortException /*ex*/)
                {
                    //MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// Set the current position of the cursor to the coordinates held in point
            /// </summary>
            /// <param name="point">Coordinates to set the cursor to</param>
            private void SetCursorPosition(Point point)
            {
                Cursor.Position = point;
            }

            #endregion
        }

        #endregion
    }
}
