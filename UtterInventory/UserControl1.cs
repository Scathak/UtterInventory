using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using AForge.Video.DirectShow;
using AForge.Video;
using MessageBox = System.Windows.Forms.MessageBox;
using ZXing;
using ZXing.Common;

namespace UtterInventory
{
    public partial class UserControl1 : System.Windows.Forms.UserControl
    {
        private VideoCapture _capture;
        private BarcodeReader _reader;
        private string previousResult = string.Empty;
        private VideoCaptureDevice videoSource;
        private string CamsProperties = string.Empty;
        private FilterInfoCollection videoDevices = null;
        private int selecteditemNumber = -1;

        public UserControl1()
        {
            InitializeComponent();
            PopulateCameras();
            InitBarcodeReader();
        }
        private void InitBarcodeReader()
        {
            _reader = new BarcodeReader
            {
                AutoRotate = true,
                Options = new DecodingOptions { TryHarder = true }
            };
        }
        private void PopulateCameras()
        {
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            var CamNum = 0;
            foreach (FilterInfo device in videoDevices)
            {
                ComboBox1.Items.Add(CamNum + ". Camera: " + device.Name);
                CamNum++;
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedCamera = ComboBox1.SelectedItem.ToString();
            if (!string.IsNullOrEmpty(selectedCamera))
            {
                if (selecteditemNumber == ComboBox1.SelectedIndex) return;
                else selecteditemNumber = ComboBox1.SelectedIndex;

                if(videoSource != null)
                {
                    videoSource.SignalToStop(); 
                }

                if (selectedCamera.Contains("http://"))
                { 
                    Globals.ThisAddIn.ipCameraAddress = selectedCamera;
                    ConnectToIPCamera();
                }
                else
                {
                    StartUSBCamera(selectedCamera);
                }
            }
        }
        private void StartUSBCamera(string cameraName)
        {
            int cameraNumber = extractNumberOfLine(cameraName);
            FilterInfo camera = videoDevices[cameraNumber];
            videoSource = new VideoCaptureDevice(camera.MonikerString);
            videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
            videoSource.Start();
        }
        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap bitmap = eventArgs.Frame;

            VideoPreviewPlay(bitmap);
            InsertDecodedQR(bitmap);
            Thread.Sleep(250);
        }
        private int extractNumberOfLine(string input) {
            int number = 0;
            int dotIndex = input.IndexOf('.');
            if (dotIndex != -1)
            {
                string numberString = input.Substring(0, dotIndex);
                if (int.TryParse(numberString, out number))
                {
                    Console.WriteLine($"The number before the full stop is: {number}");
                }
                else
                {
                    Console.WriteLine("The extracted substring is not a valid integer.");
                }
            }
            else
            {
                Console.WriteLine("No full stop found in the input string.");
            }
            return number;
        }
        public void ConnectToIPCamera()
        {
            if (!(Globals.ThisAddIn.bPlayflag))
            {
                Globals.ThisAddIn.bPlayflag = true;
                Globals.ThisAddIn.ThreadCam = new Thread(TurnOnCamera);
                Globals.ThisAddIn.ThreadCam.Start();
            }
            else
            {
                Globals.ThisAddIn.bPlayflag = false;
                Globals.ThisAddIn.ThreadCam.Abort();
            }
        }
        private void TurnOnCamera()
        {
            if (string.IsNullOrEmpty(Globals.ThisAddIn.ipCameraAddress)) return;

            // Create a VideoCapture object to get video from IP camera
            var _capture = new VideoCapture(Globals.ThisAddIn.ipCameraAddress);

            // Check if the camera opened successfully
            if (!_capture.IsOpened())
            {
                System.Windows.Forms.MessageBox.Show("Error: Unable to open the video stream from the IP camera.");
                return;
            }
            while (Globals.ThisAddIn.bPlayflag)
            {
                Mat frame = new Mat();
                // Read a frame from the video stream
                _capture.Read(frame);
                // Check if the frame is empty (end of the stream)
                if (frame.Empty())
                {
                    MessageBox.Show("Error: Unable to read a frame from the IP camera.");
                    Thread.Sleep(10);
                    break;
                }
                // Convert the frame to a format suitable for ZXing to process
                var bitmap = BitmapConverter.ToBitmap(frame);
                VideoPreviewPlay(bitmap);
                InsertDecodedQR(bitmap);

                // Dispose the bitmap after use to avoid memory leaks
                bitmap.Dispose();
                frame.Dispose();
                // Introduce a small delay to prevent overwhelming the CPU
                Thread.Sleep(100);
            }
            _capture.Release();
            Globals.ThisAddIn.ThreadCam.Abort();
        }
        private void VideoPreviewPlay(Bitmap bitmap)
        {
            // Use Invoke to safely update the PictureBox on the UI thread
            if (checkBox4.Checked)
            {
                if (pictureBoxVideo.InvokeRequired)
                {
                    pictureBoxVideo.Invoke(new Action(() => UpdatePictureBox(bitmap)));
                }
                else
                {
                    UpdatePictureBox(bitmap);
                }
            }
            else
            {
                pictureBoxVideo.Image?.Dispose();
                pictureBoxVideo.Image = null;
            }
        }
        private void InsertDecodedQR(Bitmap bitmap)
        {
            using (Bitmap clonedBitmap = (Bitmap)bitmap.Clone())
            {
                Result result = _reader.Decode(clonedBitmap);
                if (result != null)
                {
                    if (!checkBox1.Checked || result.Text != previousResult)
                    {
                        // Use Invoke to safely show the MessageBox on the UI thread
                        this.Invoke(new Action(() =>
                        {
                            System.Windows.Forms.Clipboard.SetText(result.Text);

                            if (checkBox5.Checked)
                            {
                                var sheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Paste();
                                // TODO active worksheet focus detect 
                                Thread.Sleep(10);
                            }
                            else {
                                //TODO maybe outside Invoke()
                                SendKeys.SendWait("^{v}");
                                
                                //Thread.Sleep(10);
                            }
                            previousResult = result.Text;
                            
                            if (checkBox2.Checked) SendKeys.Send("{ENTER}");
                            if (checkBox3.Checked) SendKeys.Send("{TAB}");
                            System.Windows.Forms.Clipboard.Clear();
                        }));
                        if (true) new EventSoundPlayer("chimes.wav").StartPlaySound();
                    }
                }
            }
        }
        private void UpdatePictureBox(Bitmap bitmap)
        {
            // Dispose the old image and assign the new one
            pictureBoxVideo.Image?.Dispose();
            pictureBoxVideo.Image = bitmap;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var newString = textBox1.Text;
            if (!string.IsNullOrEmpty(newString))
            {
                foreach (string element in ComboBox1.Items)
                {
                    if (element.Contains(newString)) return;
                }
                ComboBox1.Items.Add(ComboBox1.Items.Count + ". " + newString);
            }
            return;
        }


    }
}
