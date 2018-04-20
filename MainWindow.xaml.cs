using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;
using Microsoft.Speech.Recognition;
using System.Threading;
using System.IO;
using Microsoft.Speech.AudioFormat;
using System.Diagnostics;
using System.Windows.Threading;

namespace AirControl
{
    public partial class MainWindow : Window
    {
        //Initializes the Kinect Camera and the Speech Recognition engine
        KinectSensor sensor;
        SpeechRecognitionEngine speechRecognizer;

        //Starts the 'real clock' timer used in measuring fps
        DispatcherTimer readyTimer;

        //Initializes the byte array
        byte[] colorBytes;

        //The skeletal framework the Kinect offers was used
        Skeleton[] skeletons;

        // Definition of important boolean variables to be used later
        bool isCirclesVisible = true;

        bool isForwardGestureActive = false;
        bool isBackGestureActive = false;
        bool isCrossGestureActive = false;
        bool isPushGestureActive = false;
        bool isPullGestureActive = false;
        bool isLaunchSlideActive = false;

        // Definition of color for when commands are being executed or not
        SolidColorBrush activeBrush = new SolidColorBrush(Colors.LightSkyBlue);
        SolidColorBrush inactiveBrush = new SolidColorBrush(Colors.LightGray);

        public MainWindow()
        {
            InitializeComponent();

            //Runtime initialization is handled when the window is opened. When the window
            //is closed, the runtime MUST be unitialized.
            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
            //Handles the content obtained from the video camera, once received.

            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);
        }

        // Test the proper Pc to Kinect conection
        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            sensor = KinectSensor.KinectSensors.FirstOrDefault();

            if (sensor == null)
            {
                MessageBox.Show("Please properly configure the Kinect Sensor.");
                this.Close();
            }

            // Initalize live camera feed from the sensor
            sensor.Start();

            // Standard resolution for video and depth detection
            sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
            sensor.ColorFrameReady += new EventHandler<ColorImageFrameReadyEventArgs>(sensor_ColorFrameReady);

            sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);

            sensor.SkeletonStream.Enable();
            sensor.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(sensor_SkeletonFrameReady);

            // Edits the sensor's elevation angle. Moslty not needed, but just in case

            sensor.ElevationAngle = 20;


            Application.Current.Exit += new ExitEventHandler(Current_Exit);

            //Set up simple speech commands
            InitializeSpeechRecognition();
        }

        void Current_Exit(object sender, ExitEventArgs e)
        {
            // Test for Microsoft Speech compatability
            if (speechRecognizer != null)
            {
                speechRecognizer.RecognizeAsyncCancel();
                speechRecognizer.RecognizeAsyncStop();
            }
            if (sensor != null)
            {
                sensor.AudioSource.Stop();
                sensor.Stop();
                sensor.Dispose();
                sensor = null;
            }
        }

        void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            //Toggle the display of the command display circles
            if (e.Key == Key.C)
            {
                ToggleCircles();
            }
        }
        
        // This tells the Kinect that we want complete RGB live feed 
        void sensor_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (var image = e.OpenColorImageFrame())
            {
                if (image == null)
                    return;

                if (colorBytes == null ||
                    colorBytes.Length != image.PixelDataLength)
                {
                    colorBytes = new byte[image.PixelDataLength];
                }

                image.CopyPixelDataTo(colorBytes);

                /*You could use PixelFormats.Bgr32 below to ignore the alpha,
                or if you need to set the alpha you would loop through the bytes 
                as in this loop below*/

                int length = colorBytes.Length;
                for (int i = 0; i < length; i += 4)
                {
                    colorBytes[i + 3] = 255;
                }

                // Dictate the quality of the sensor detection through bitmapping
                BitmapSource source = BitmapSource.Create(image.Width,
                    image.Height,
                    96,
                    96,
                    PixelFormats.Bgra32,
                    null,
                    colorBytes,
                    image.Width * image.BytesPerPixel);
                videoImage.Source = source;
            }
        }
        
        // Prepare the initialized skeleton framework for its specific use
        void sensor_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            using (var skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame == null)
                    return;

                if (skeletons == null ||
                    skeletons.Length != skeletonFrame.SkeletonArrayLength)
                {
                    skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                }

                skeletonFrame.CopySkeletonDataTo(skeletons);
            }

            Skeleton closestSkeleton = skeletons.Where(s => s.TrackingState == SkeletonTrackingState.Tracked)
                                                .OrderBy(s => s.Position.Z * Math.Abs(s.Position.X))
                                                .FirstOrDefault();

            if (closestSkeleton == null)
                return;

            // The only three parts of the skeletal framework used
            var head = closestSkeleton.Joints[JointType.Head];
            var rightHand = closestSkeleton.Joints[JointType.HandRight];
            var leftHand = closestSkeleton.Joints[JointType.HandLeft];

            // This turns of the [circle]  trackers if one of the body parts cannot be identified
            if (head.TrackingState == JointTrackingState.NotTracked ||
                rightHand.TrackingState == JointTrackingState.NotTracked ||
                leftHand.TrackingState == JointTrackingState.NotTracked)
            {
                
                return;
            }

            // Set the possible states for each of the three circles
            SetEllipsePosition(ellipseHead, head, false);
            SetEllipsePosition(ellipseLeftHand, leftHand, isBackGestureActive || isCrossGestureActive || isPushGestureActive || isLaunchSlideActive);
            SetEllipsePosition(ellipseRightHand, rightHand, isForwardGestureActive || isCrossGestureActive || isPullGestureActive || isLaunchSlideActive);
            

            ProcessForwardBackGesture(head, rightHand, leftHand);
        }

        //This method shows exactly when a command is being executed (red turns green)
        private void SetEllipsePosition(Ellipse ellipse, Joint joint, bool isHighlighted)
        {
            if (isHighlighted)
            {
                ellipse.Width = 20;
                ellipse.Height = 20;
                ellipse.Fill = activeBrush;
            }
            else
            {
                ellipse.Width = 20;
                ellipse.Height = 20;
                ellipse.Fill = inactiveBrush;
            }

            // Final set up for RGB feed and skeletal framework to scale the sensors for gestures
            CoordinateMapper mapper = sensor.CoordinateMapper;

            var point = mapper.MapSkeletonPointToColorPoint(joint.Position, sensor.ColorStream.Format);

            Canvas.SetLeft(ellipse, point.X - ellipse.ActualWidth / 2);
            Canvas.SetTop(ellipse, point.Y - ellipse.ActualHeight / 2);
        }
        
        // These boolean arguments provide the basis for the commands to be executed
        private void ProcessForwardBackGesture(Joint head, Joint rightHand, Joint leftHand)
        {   
            // Hand gestures are relative to the position of the head

            // If the right hand is extended to the right, execute the right arrow key
            if (rightHand.Position.X > head.Position.X + 0.45)
            {
                if (!isForwardGestureActive)
                {
                    isForwardGestureActive = true;
                    System.Windows.Forms.SendKeys.SendWait("{Right}"); // The keyboard key right is executed in the active application
                }
            }
            else
            {
                isForwardGestureActive = false; //prevents repititive execution of the arrow key
            }

            // If the left hand is extended to the left, execute thr left arrow key
            if (leftHand.Position.X < head.Position.X - 0.45)
            {
                if (!isBackGestureActive)
                {
                    isBackGestureActive = true;
                    System.Windows.Forms.SendKeys.SendWait("{Left}");
                }
            }
            else
            {
                isBackGestureActive = false;
            }

            if (rightHand.Position.X < leftHand.Position.X - 0.20) 
            {
                if (!isCrossGestureActive)
                {
                    isCrossGestureActive = true;
                    System.Windows.Forms.SendKeys.SendWait("{ESC}");
                }
                
            }
            else
            {
                isCrossGestureActive = false;
            }

            if (rightHand.Position.Y > head.Position.Y + 0.20)
            {
                if (!isLaunchSlideActive)
                {
                    isLaunchSlideActive = true;
                    System.Windows.Forms.SendKeys.SendWait("{F5}");
                }

            }
            else
            {
                isLaunchSlideActive = false;
            }

            if (leftHand.Position.Y > head.Position.Y + 0.20)
            {
                if (!isLaunchSlideActive)
                {
                    isLaunchSlideActive = true;
                    System.Windows.Forms.SendKeys.SendWait("{F5}");
                }

            }
            else
            {
                isLaunchSlideActive = false;
            }


            // Using 'the force' to control the enter and backspace keys
            if (leftHand.Position.Z < head.Position.Z - 0.45)
            {
                if(!isPushGestureActive)
                {
                    isPushGestureActive = true;
                    System.Windows.Forms.SendKeys.SendWait("{Up}");
                }

                else
                {
                    isPushGestureActive = false;
                }
            }

            //When using the z-coordinate, the commands act differently. It repeatedly loops the command about 3 times a second, which is very queer but has no definite solution.
            if (rightHand.Position.Z < head.Position.Z - 0.45)
            {
                if (!isPullGestureActive)
                {
                    isPullGestureActive = true;
                    System.Windows.Forms.SendKeys.SendWait("{Down}");
                }

                else
                {
                    isPullGestureActive = false;
                }
            }

        }
        

        // Code to toggle the visibility of the circles
        void ToggleCircles()
        {
            if (isCirclesVisible)
                HideCircles();
            else
                ShowCircles();
        }

        void HideCircles()
        {
            isCirclesVisible = false;
            ellipseHead.Visibility = System.Windows.Visibility.Collapsed;
            ellipseLeftHand.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightHand.Visibility = System.Windows.Visibility.Collapsed;
        }

        void ShowCircles()
        {
            isCirclesVisible = true;
            ellipseHead.Visibility = System.Windows.Visibility.Visible;
            ellipseLeftHand.Visibility = System.Windows.Visibility.Visible;
            ellipseRightHand.Visibility = System.Windows.Visibility.Visible;
        }
        // External Commands
        void Shoot()
        {
            System.Windows.Forms.SendKeys.SendWait("{Enter}");
            
        }
        void Paragraph()
        {
            System.Windows.Forms.SendKeys.SendWait("{TAB}");

        }
        void Start()
        {
            System.Windows.Forms.SendKeys.SendWait("^{ESC}");

        }
        void Switch()
        {
            System.Windows.Forms.SendKeys.SendWait("%{TAB}");

        }
        void Closer()
        {
            System.Windows.Forms.SendKeys.SendWait("%{F4}");

        }
        void Screenshot()
        {
            System.Windows.Forms.SendKeys.SendWait("{PRTSC}");

        }


        // Code to toggle the visibility of the app window
        private void ShowWindow()
        {
            this.Topmost = true;
            this.WindowState = System.Windows.WindowState.Maximized;
        }

        private void HideWindow()
        {
            this.Topmost = false;
            this.WindowState = System.Windows.WindowState.Minimized;
        }

       
       
        #region Speech Recognition Methods

        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }


        // Make sure it is known to the user that the Speech recongition does not work, if it doesn't.
        private void InitializeSpeechRecognition()
        {
            RecognizerInfo ri = GetKinectRecognizer();
            if (ri == null)
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            try
            {
                speechRecognizer = new SpeechRecognitionEngine(ri.Id);
            }
            catch
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed and configured.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }

            var phrases = new Choices();
            phrases.Add("computer show window");
            phrases.Add("computer hide window");
            phrases.Add("computer show circles");
            phrases.Add("computer hide circles");
            phrases.Add("shoot");
            phrases.Add("paragraph");
            phrases.Add("start");
            phrases.Add("switch");
            phrases.Add("close");
            phrases.Add("screenshot");

            var gb = new GrammarBuilder();
            //Specify the culture to match the recognizer in case we are running in a different culture.                                 
            gb.Culture = ri.Culture;
            gb.Append(phrases);

            // Create the actual Grammar instance, and then load it into the speech recognizer.
            var g = new Grammar(gb);

            speechRecognizer.LoadGrammar(g);
            speechRecognizer.SpeechRecognized += SreSpeechRecognized;
            speechRecognizer.SpeechHypothesized += SreSpeechHypothesized;
            speechRecognizer.SpeechRecognitionRejected += SreSpeechRecognitionRejected;

            this.readyTimer = new DispatcherTimer();
            this.readyTimer.Tick += this.ReadyTimerTick;
            this.readyTimer.Interval = new TimeSpan(0, 0, 4);
            this.readyTimer.Start();

        }

        private void ReadyTimerTick(object sender, EventArgs e)
        {
            this.StartSpeechRecognition();
            this.readyTimer.Stop();
            this.readyTimer.Tick -= ReadyTimerTick;
            this.readyTimer = null;
        }

        private void StartSpeechRecognition()
        {
            if (sensor == null || speechRecognizer == null)
                return;

            var audioSource = this.sensor.AudioSource;
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            var kinectStream = audioSource.Start();

            speechRecognizer.SetInputToAudioStream(
                    kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);

        }

        void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            Trace.WriteLine("\nSpeech Rejected, confidence: " + e.Result.Confidence);
        }

        void SreSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            Trace.Write("\rSpeech Hypothesized: \t{0}", e.Result.Text);
        }

        void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            /*This release of the Kinect language pack doesn't have a  too-reliable confidence model, so 
            we don't use e.Result.Confidence here. */

            if (e.Result.Confidence < 0.40)
            {
                Trace.WriteLine("\nSpeech Rejected filtered, confidence: " + e.Result.Confidence);
                return;
            }

            Trace.WriteLine("\nSpeech Recognized, confidence: " + e.Result.Confidence + ": \t{0}", e.Result.Text);

            if (e.Result.Text == "computer show window")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                    {
                        ShowWindow();
                    });
            }
            else if (e.Result.Text == "computer hide window")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    HideWindow();
                });
            }
            else if (e.Result.Text == "computer hide circles")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.HideCircles();
                });
            }
            else if (e.Result.Text == "computer show circles")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.ShowCircles();
                });
            }
            else if (e.Result.Text == "shoot")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.Shoot();
                });
            }
            else if (e.Result.Text == "paragraph")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.Paragraph();
                });
            }
            else if (e.Result.Text == "start")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.Start();
                });
            }
            else if (e.Result.Text == "switch")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.Switch();
                });
            }
            else if (e.Result.Text == "close")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.Closer();
                });
            }
            else if (e.Result.Text == "screenshot")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.Screenshot();
                });
            }
        }

        #endregion

    }
}
