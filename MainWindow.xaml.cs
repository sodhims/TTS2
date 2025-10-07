using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;  // This is needed for KeyEventArgs
using System.Windows.Media;
using Microsoft.Win32;
using NAudio.Wave;
using NAudio.Lame;
using System.Net.Http;
using System.Text.Json;
using Amazon;
using Amazon.Polly;
using Amazon.Polly.Model;
using Amazon.Runtime;
using System.Configuration;


namespace TTS2
{
    public partial class MainWindow : Window
    {
        private SpeechSynthesizer speechSynthesizer;
        private HttpClient httpClient;
        private string googleApiKey = "";
        private string currentFilePath = "";
        private bool isPlaying = false;
        private WaveOutEvent waveOut;
        private AudioFileReader currentAudioFile;
        private string currentGoogleVoice = "en-US-Wavenet-D";
        private List<string> createdAudioFiles = new List<string>();
        private int currentPlayingIndex = -1;
        private System.Windows.Threading.DispatcherTimer playbackTimer;
        private string lastOutputFolder = "";
        private AmazonPollyClient pollyClient;
        private string awsAccessKey = "";
        private string awsSecretKey = "";
        private string awsRegion = "us-east-1";
        private string currentAwsVoice = "Joanna";
        private bool useAwsNeuralEngine = false;

        private readonly Dictionary<string, VoiceInfo> awsVoices = new Dictionary<string, VoiceInfo>
        {
            // Standard US English voices
            { "Joanna (Female, US)", new VoiceInfo { VoiceId = "Joanna", Engine = "standard" } },
            { "Matthew (Male, US)", new VoiceInfo { VoiceId = "Matthew", Engine = "standard" } },
            { "Ivy (Female, Child, US)", new VoiceInfo { VoiceId = "Ivy", Engine = "standard" } },
            { "Joey (Male, US)", new VoiceInfo { VoiceId = "Joey", Engine = "standard" } },
            { "Justin (Male, Child, US)", new VoiceInfo { VoiceId = "Justin", Engine = "standard" } },
            { "Kendra (Female, US)", new VoiceInfo { VoiceId = "Kendra", Engine = "standard" } },
            { "Kimberly (Female, US)", new VoiceInfo { VoiceId = "Kimberly", Engine = "standard" } },
            { "Salli (Female, US)", new VoiceInfo { VoiceId = "Salli", Engine = "standard" } },
            
            // Neural US English voices
            { "Joanna Neural (Female, US)", new VoiceInfo { VoiceId = "Joanna", Engine = "neural" } },
            { "Matthew Neural (Male, US)", new VoiceInfo { VoiceId = "Matthew", Engine = "neural" } },
            { "Kevin Neural (Male, US)", new VoiceInfo { VoiceId = "Kevin", Engine = "neural" } },
            { "Ruth Neural (Female, US)", new VoiceInfo { VoiceId = "Ruth", Engine = "neural" } },
            { "Ivy Neural (Female, Child, US)", new VoiceInfo { VoiceId = "Ivy", Engine = "neural" } },
            { "Joey Neural (Male, US)", new VoiceInfo { VoiceId = "Joey", Engine = "neural" } },
            { "Justin Neural (Male, Child, US)", new VoiceInfo { VoiceId = "Justin", Engine = "neural" } },
            { "Kendra Neural (Female, US)", new VoiceInfo { VoiceId = "Kendra", Engine = "neural" } },
            { "Kimberly Neural (Female, US)", new VoiceInfo { VoiceId = "Kimberly", Engine = "neural" } },
            { "Salli Neural (Female, US)", new VoiceInfo { VoiceId = "Salli", Engine = "neural" } },
            
            // Other English variants
            { "Nicole (Female, AU)", new VoiceInfo { VoiceId = "Nicole", Engine = "standard" } },
            { "Russell (Male, AU)", new VoiceInfo { VoiceId = "Russell", Engine = "standard" } },
            { "Amy (Female, GB)", new VoiceInfo { VoiceId = "Amy", Engine = "standard" } },
            { "Brian (Male, GB)", new VoiceInfo { VoiceId = "Brian", Engine = "standard" } },
            { "Emma (Female, GB)", new VoiceInfo { VoiceId = "Emma", Engine = "standard" } },
            { "Amy Neural (Female, GB)", new VoiceInfo { VoiceId = "Amy", Engine = "neural" } },
            { "Brian Neural (Male, GB)", new VoiceInfo { VoiceId = "Brian", Engine = "neural" } },
            { "Emma Neural (Female, GB)", new VoiceInfo { VoiceId = "Emma", Engine = "neural" } },
            { "Aditi (Female, IN)", new VoiceInfo { VoiceId = "Aditi", Engine = "standard" } },
            { "Raveena (Female, IN)", new VoiceInfo { VoiceId = "Raveena", Engine = "standard" } }
        };

        // Helper class for AWS voice information
        public class VoiceInfo
        {
            public string VoiceId { get; set; }
            public string Engine { get; set; }
        }
        
        private readonly Dictionary<string, string> googleVoices = new Dictionary<string, string>
        {
            { "en-US-Wavenet-A (Female)", "en-US-Wavenet-A" },
            { "en-US-Wavenet-B (Male)", "en-US-Wavenet-B" },
            { "en-US-Wavenet-C (Female)", "en-US-Wavenet-C" },
            { "en-US-Wavenet-D (Male)", "en-US-Wavenet-D" },
            { "en-US-Wavenet-E (Female)", "en-US-Wavenet-E" },
            { "en-US-Wavenet-F (Female)", "en-US-Wavenet-F" },
            { "en-US-Neural2-A (Male)", "en-US-Neural2-A" },
            { "en-US-Neural2-C (Female)", "en-US-Neural2-C" },
            { "en-US-Neural2-D (Male)", "en-US-Neural2-D" },
            { "en-US-Neural2-E (Female)", "en-US-Neural2-E" },
            { "en-US-Standard-A (Male)", "en-US-Standard-A" },
            { "en-US-Standard-B (Male)", "en-US-Standard-B" },
            { "en-US-Standard-C (Female)", "en-US-Standard-C" },
            { "en-US-Standard-D (Male)", "en-US-Standard-D" },
            { "en-US-Standard-E (Female)", "en-US-Standard-E" }
        };
        private string elevenLabsApiKey = "";
        private string currentElevenLabsVoice = "21m00Tcm4TlvDq8ikWAM"; // Rachel voice ID
        private readonly Dictionary<string, string> elevenLabsVoices = new Dictionary<string, string>
        {
            // Pre-made voices (free tier)
            { "Rachel (Female, US)", "21m00Tcm4TlvDq8ikWAM" },
            { "Domi (Female, US)", "AZnzlk1XvdvUeBnXmlld" },
            { "Bella (Female, US)", "EXAVITQu4vr4xnSDxMaL" },
            { "Antoni (Male, US)", "ErXwobaYiN019PkySvjV" },
            { "Elli (Female, US)", "MF3mGyEYCl7XYWbV9V6O" },
            { "Josh (Male, US)", "TxGEqnHWrfWFTfGW9XjX" },
            { "Arnold (Male, US)", "VR6AewLTigWG4xSOukaG" },
            { "Adam (Male, US)", "pNInz6obpgDQGcFmaJgB" },
            { "Sam (Male, US)", "yoZ06aMxZJJ28mfd3POQ" },
            // Add your custom cloned voices here
        };
        public MainWindow()
        {
            InitializeComponent();
            InitializeTTS();
            httpClient = new HttpClient();
            httpClient.Timeout = TimeSpan.FromSeconds(30);
            cmbTTSEngine.SelectedIndex = 0;
            cmbOutputFormat.SelectedIndex = 0;
            LoadVoices();
            
            // Initialize playback timer
            playbackTimer = new System.Windows.Threading.DispatcherTimer();
            playbackTimer.Interval = TimeSpan.FromMilliseconds(100);
            playbackTimer.Tick += PlaybackTimer_Tick;
			this.PreviewKeyDown += MainWindow_PreviewKeyDown;

            //Load saved credentials
            LoadCredentials();
        }
        private void LoadCredentials()
        {
            try
            {
                awsAccessKey = ConfigurationManager.AppSettings["AWSAccessKey"] ?? "";
                awsSecretKey = ConfigurationManager.AppSettings["AWSSecretKey"] ?? "";
                awsRegion = ConfigurationManager.AppSettings["AWSRegion"] ?? "us-east-1";
                googleApiKey = ConfigurationManager.AppSettings["GoogleApiKey"] ?? "";
                elevenLabsApiKey = ConfigurationManager.AppSettings["ElevenLabsApiKey"] ?? "";
                
                // Initialize AWS client if credentials exist
                if (!string.IsNullOrEmpty(awsAccessKey) && !string.IsNullOrEmpty(awsSecretKey))
                {
                    var credentials = new BasicAWSCredentials(awsAccessKey, awsSecretKey);
                    var region = RegionEndpoint.GetBySystemName(awsRegion);
                    pollyClient = new AmazonPollyClient(credentials, region);
                    LogMessage($"AWS credentials loaded from config (Region: {awsRegion})");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error loading credentials: {ex.Message}");
            }
        }

        private void SaveCredentials()
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                
                config.AppSettings.Settings.Remove("AWSAccessKey");
                config.AppSettings.Settings.Add("AWSAccessKey", awsAccessKey);
                
                config.AppSettings.Settings.Remove("AWSSecretKey");
                config.AppSettings.Settings.Add("AWSSecretKey", awsSecretKey);
                
                config.AppSettings.Settings.Remove("AWSRegion");
                config.AppSettings.Settings.Add("AWSRegion", awsRegion);
                
                config.AppSettings.Settings.Remove("GoogleApiKey");
                config.AppSettings.Settings.Add("GoogleApiKey", googleApiKey);
                

                config.AppSettings.Settings.Remove("ElevenLabsApiKey");
                config.AppSettings.Settings.Add("ElevenLabsApiKey", elevenLabsApiKey);

                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");

                LogMessage("Credentials saved to config file");
            }
            catch (Exception ex)
            {
                LogMessage($"Error saving credentials: {ex.Message}");
            }
        }
		

        
        private void InitializeTTS()
        {
            speechSynthesizer = new SpeechSynthesizer();
            speechSynthesizer.SetOutputToDefaultAudioDevice();
        }
        
        private void LoadVoices()
        {
            cmbVoices.Items.Clear();
            
            if (cmbTTSEngine.SelectedIndex == 0) // Windows SAPI
            {
                var voices = speechSynthesizer.GetInstalledVoices();
                foreach (var voice in voices)
                {
                    cmbVoices.Items.Add(voice.VoiceInfo.Name);
                }
            }
            else if (cmbTTSEngine.SelectedIndex == 1) // Google TTS
            {
                foreach (var voice in googleVoices.Keys)
                {
                    cmbVoices.Items.Add(voice);
                }
            }
            else if (cmbTTSEngine.SelectedIndex == 2) // AWS Polly
            {
                foreach (var voice in awsVoices.Keys)
                {
                    cmbVoices.Items.Add(voice);
                }
            }
            else if (cmbTTSEngine.SelectedIndex == 3) // ElevenLabs
            {
                foreach (var voice in elevenLabsVoices.Keys)
                {
                    cmbVoices.Items.Add(voice);
                }
                        if (cmbVoices.Items.Count > 0)
                            cmbVoices.SelectedIndex = 0;
            }
        }
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            };
            
            if (openFileDialog.ShowDialog() == true)
            {
                currentFilePath = openFileDialog.FileName;
                string content = File.ReadAllText(currentFilePath);
                
                rtbTextContent.Document.Blocks.Clear();
                rtbTextContent.Document.Blocks.Add(new Paragraph(new Run(content)));
                
                LogMessage($"Loaded file: {currentFilePath}");
            }
        }
        
        private void SaveText_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            };
            
            if (saveFileDialog.ShowDialog() == true)
            {
                string text = new TextRange(rtbTextContent.Document.ContentStart, 
                    rtbTextContent.Document.ContentEnd).Text;
                File.WriteAllText(saveFileDialog.FileName, text);
                LogMessage($"Saved file: {saveFileDialog.FileName}");
            }
        }
        
        private void TTSEngine_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (btnAPIKey != null && btnAWSCredentials != null)
            {
                // Show appropriate credential button based on selected engine
                btnAPIKey.Visibility = cmbTTSEngine.SelectedIndex == 1 ? 
                    Visibility.Visible : Visibility.Collapsed;
                
                btnAWSCredentials.Visibility = cmbTTSEngine.SelectedIndex == 2 ? 
                    Visibility.Visible : Visibility.Collapsed;

                btnElevenLabsKey.Visibility = cmbTTSEngine.SelectedIndex == 3 ? 
                    Visibility.Visible : Visibility.Collapsed;
            }
            LoadVoices();
        }
        
        // private void APIKey_Click(object sender, RoutedEventArgs e)
        // {
            // var dialog = new Window
            // {
                // Title = "Enter Google Cloud API Key",
                // Width = 400,
                // Height = 150,
                // WindowStartupLocation = WindowStartupLocation.CenterOwner,
                // Owner = this
            // };
            
            // var grid = new Grid();
            // var textBox = new TextBox 
            // { 
                // Margin = new Thickness(10),
                // VerticalAlignment = VerticalAlignment.Center,
                // Text = googleApiKey
            // };
            
            // var buttonPanel = new StackPanel
            // {
                // Orientation = Orientation.Horizontal,
                // HorizontalAlignment = HorizontalAlignment.Right,
                // VerticalAlignment = VerticalAlignment.Bottom,
                // Margin = new Thickness(10)
            // };
            
            // var okButton = new Button { Content = "OK", Width = 75, Margin = new Thickness(5) };
            // var cancelButton = new Button { Content = "Cancel", Width = 75, Margin = new Thickness(5) };
            
            // okButton.Click += (s, args) =>
            // {
                // googleApiKey = textBox.Text;
                // LogMessage("Google API Key set successfully");
                // dialog.DialogResult = true;
            // };
            
            // cancelButton.Click += (s, args) => dialog.DialogResult = false;
            
            // buttonPanel.Children.Add(okButton);
            // buttonPanel.Children.Add(cancelButton);
            
            // grid.Children.Add(textBox);
            // grid.Children.Add(buttonPanel);
            // dialog.Content = grid;
            
            // dialog.ShowDialog();
        // }
        
        private void InsertSplit_Click(object sender, RoutedEventArgs e)
        {
            var caretPos = rtbTextContent.CaretPosition;
            caretPos.InsertTextInRun("<split>");
        }
        
        private void InsertVoice_Click(object sender, RoutedEventArgs e)
        {

            if (cmbVoiceTag.SelectedItem != null)
            {
                var voiceTag = (cmbVoiceTag.SelectedItem as ComboBoxItem).Content.ToString();
                var caretPos = rtbTextContent.CaretPosition;
                caretPos.InsertTextInRun($"<{voiceTag}>");
            }
        }
        
        private void InsertSSML_Click(object sender, RoutedEventArgs e)
        {
            if (cmbSSMLTags.SelectedItem != null)
            {
                var ssmlTag = (cmbSSMLTags.SelectedItem as ComboBoxItem).Content.ToString();
                var caretPos = rtbTextContent.CaretPosition;
                
                // Check if it's a self-closing tag
                if (ssmlTag.Contains("/>"))
                {
                    caretPos.InsertTextInRun(ssmlTag);
                }
                else
                {
                    // Extract tag name and insert with cursor positioned between tags
                    var tagMatch = Regex.Match(ssmlTag, @"<(\w+)[^>]*>");
                    if (tagMatch.Success)
                    {
                        var tagName = tagMatch.Groups[1].Value;
                        var openTag = tagMatch.Groups[0].Value;
                        var closeTag = $"</{tagName}>";
                        
                        caretPos.InsertTextInRun(openTag);
                        var middlePos = caretPos.GetPositionAtOffset(0);
                        caretPos.InsertTextInRun(closeTag);
                        
                        // Position cursor between tags
                        rtbTextContent.CaretPosition = middlePos;
                    }
                    else
                    {
                        caretPos.InsertTextInRun(ssmlTag);
                    }
                }
            }
        }
        
        private void WrapSSML_Click(object sender, RoutedEventArgs e)
        {
            if (cmbSSMLTags.SelectedItem == null) return;
            
            var selection = rtbTextContent.Selection;
            if (!selection.IsEmpty)
            {
                string selectedText = selection.Text;
                var ssmlTag = (cmbSSMLTags.SelectedItem as ComboBoxItem).Content.ToString();
                
                // Don't wrap with self-closing tags
                if (ssmlTag.Contains("/>"))
                {
                    MessageBox.Show("Cannot wrap text with a self-closing tag. Please select a tag with opening and closing elements.");
                    return;
                }
                
                // Extract tag name and wrap selected text
                var tagMatch = Regex.Match(ssmlTag, @"<(\w+)[^>]*>");
                if (tagMatch.Success)
                {
                    var tagName = tagMatch.Groups[1].Value;
                    var openTag = tagMatch.Groups[0].Value;
                    var closeTag = $"</{tagName}>";
                    
                    string wrappedText = openTag + selectedText + closeTag;
                    selection.Text = wrappedText;
                    
                    LogMessage($"Wrapped selection with {tagName} tags");
                }
            }
            else
            {
                MessageBox.Show("Please select text to wrap with SSML tags.");
            }
        }
        
        private void Voice_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (cmbTTSEngine.SelectedIndex == 0 && cmbVoices.SelectedItem != null)
            {
                speechSynthesizer.SelectVoice(cmbVoices.SelectedItem.ToString());
            }
            else if (cmbTTSEngine.SelectedIndex == 1 && cmbVoices.SelectedItem != null)
            {
                string selectedVoice = cmbVoices.SelectedItem.ToString();
                if (googleVoices.ContainsKey(selectedVoice))
                {
                    currentGoogleVoice = googleVoices[selectedVoice];
                }
            }
            else if (cmbTTSEngine.SelectedIndex == 2 && cmbVoices.SelectedItem != null)
            {
                string selectedVoice = cmbVoices.SelectedItem.ToString();
                if (awsVoices.ContainsKey(selectedVoice))
                {
                    currentAwsVoice = awsVoices[selectedVoice].VoiceId;
                    useAwsNeuralEngine = awsVoices[selectedVoice].Engine == "neural";
                }
            }
            else if (cmbTTSEngine.SelectedIndex == 3 && cmbVoices.SelectedItem != null)
            {
                string selectedVoice = cmbVoices.SelectedItem.ToString();
                if (elevenLabsVoices.ContainsKey(selectedVoice))
                {
                    currentElevenLabsVoice = elevenLabsVoices[selectedVoice];
                }
            }
        }
        private void AWSCredentials_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Window
            {
                Title = "AWS Polly Credentials",
                Width = 500,
                Height = 300,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize
            };
            
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            
            // Access Key
            var accessKeyLabel = new TextBlock
            {
                Text = "AWS Access Key ID:",
                Margin = new Thickness(10, 10, 10, 5),
                FontWeight = FontWeights.SemiBold
            };
            
            var accessKeyBox = new TextBox 
            { 
                Margin = new Thickness(10, 0, 10, 10),
                Text = awsAccessKey,
                FontFamily = new FontFamily("Consolas"),
                Height = 25
            };
            
            // Secret Key
            var secretKeyLabel = new TextBlock
            {
                Text = "AWS Secret Access Key:",
                Margin = new Thickness(10, 0, 10, 5),
                FontWeight = FontWeights.SemiBold
            };
            
            var secretKeyBox = new PasswordBox 
            { 
                Margin = new Thickness(10, 0, 10, 10),
                FontFamily = new FontFamily("Consolas"),
                Height = 25
            };
            
            if (!string.IsNullOrEmpty(awsSecretKey))
            {
                secretKeyBox.Password = awsSecretKey;
            }
            
            // Region
            var regionLabel = new TextBlock
            {
                Text = "AWS Region:",
                Margin = new Thickness(10, 0, 10, 5),
                FontWeight = FontWeights.SemiBold
            };
            
            var regionCombo = new ComboBox
            {
                Margin = new Thickness(10, 0, 10, 10),
                Height = 25
            };
            
            var regions = new[] 
            { 
                "us-east-1", "us-west-2", "us-west-1", "eu-west-1", 
                "eu-central-1", "ap-southeast-1", "ap-northeast-1", 
                "ap-southeast-2", "ap-south-1", "sa-east-1" 
            };
            
            foreach (var region in regions)
            {
                regionCombo.Items.Add(region);
            }
            
            regionCombo.SelectedItem = awsRegion;
            if (regionCombo.SelectedItem == null && regionCombo.Items.Count > 0)
            {
                regionCombo.SelectedIndex = 0;
            }
            
            // Instructions
            var instructionText = new TextBlock
            {
                Text = "Get credentials from AWS IAM Console.\nEnsure your IAM user has AmazonPollyFullAccess policy.",
                Margin = new Thickness(10, 0, 10, 10),
                FontSize = 11,
                Foreground = Brushes.Gray,
                TextWrapping = TextWrapping.Wrap
            };
            
            // Buttons
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(10)
            };
            
            var okButton = new Button { Content = "OK", Width = 75, Margin = new Thickness(5) };
            var cancelButton = new Button { Content = "Cancel", Width = 75, Margin = new Thickness(5) };
            
            okButton.Click += (s, args) =>
            {
                awsAccessKey = accessKeyBox.Text.Trim();
                awsSecretKey = secretKeyBox.Password.Trim();
                awsRegion = regionCombo.SelectedItem?.ToString() ?? "us-east-1";
                
                if (string.IsNullOrEmpty(awsAccessKey) || string.IsNullOrEmpty(awsSecretKey))
                {
                    LogMessage("AWS credentials cleared");
                    if (pollyClient != null)
                    {
                        pollyClient.Dispose();
                        pollyClient = null;
                    }
                }
                else
                {
                    try
                    {
                        var credentials = new BasicAWSCredentials(awsAccessKey, awsSecretKey);
                        var region = RegionEndpoint.GetBySystemName(awsRegion);
                        pollyClient = new AmazonPollyClient(credentials, region);
                        LogMessage($"AWS Polly credentials set successfully for region: {awsRegion}");
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Error initializing AWS Polly: {ex.Message}");
                        MessageBox.Show($"Error initializing AWS Polly: {ex.Message}", 
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                
                // Save credentials
                SaveCredentials();
                
                dialog.DialogResult = true;
            };
            
            cancelButton.Click += (s, args) => dialog.DialogResult = false;
            
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            
            Grid.SetRow(accessKeyLabel, 0);
            Grid.SetRow(accessKeyBox, 1);
            Grid.SetRow(secretKeyLabel, 2);
            Grid.SetRow(secretKeyBox, 3);
            Grid.SetRow(regionLabel, 4);
            Grid.SetRow(regionCombo, 5);
            Grid.SetRow(instructionText, 6);
            Grid.SetRow(buttonPanel, 7);
            
            grid.Children.Add(accessKeyLabel);
            grid.Children.Add(accessKeyBox);
            grid.Children.Add(secretKeyLabel);
            grid.Children.Add(secretKeyBox);
            grid.Children.Add(regionLabel);
            grid.Children.Add(regionCombo);
            grid.Children.Add(instructionText);
            grid.Children.Add(buttonPanel);
            
            dialog.Content = grid;
            
            dialog.Loaded += (s, args) => 
            {
                accessKeyBox.Focus();
                accessKeyBox.SelectAll();
            };
            
            dialog.ShowDialog();
        }

        
        private async void TestVoice_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string testText = "Hello, this is a test of the selected voice.";
                
                if (cmbTTSEngine.SelectedIndex == 0) // Windows SAPI
                {
                    speechSynthesizer.Rate = (int)sliderRate.Value;
                    speechSynthesizer.Volume = (int)sliderVolume.Value;
                    speechSynthesizer.SpeakAsync(testText);
                    isPlaying = true;
                }
                else if (cmbTTSEngine.SelectedIndex == 1) // Google TTS
                {
                    if (string.IsNullOrEmpty(googleApiKey))
                    {
                        MessageBox.Show("Please set Google API key first");
                        return;
                    }
                    
                    string tempFile1 = Path.GetTempFileName() + ".wav";
                    bool success1 = await CallGoogleTTS(testText, tempFile1);
                    
                    if (success1 && File.Exists(tempFile1))
                    {
                        waveOut = new WaveOutEvent();
                        var audioFile = new AudioFileReader(tempFile1);
                        waveOut.Init(audioFile);
                        waveOut.Play();
                        waveOut.PlaybackStopped += (s, args) =>
                        {
                            audioFile.Dispose();
                            File.Delete(tempFile1);
                        };
                    }
                }
                else if (cmbTTSEngine.SelectedIndex == 2) // AWS Polly
                {
                    if (pollyClient == null)
                    {
                        MessageBox.Show("Please set AWS credentials first");
                        return;
                    }
                    
                    string tempFile2 = Path.GetTempFileName() + ".wav";
                    bool success2 = await CallAWSPolly(testText, tempFile2);
                    
                    if (success2 && File.Exists(tempFile2))
                    {
                        waveOut = new WaveOutEvent();
                        var audioFile = new AudioFileReader(tempFile2);
                        waveOut.Init(audioFile);
                        waveOut.Play();
                        waveOut.PlaybackStopped += (s, args) =>
                        {
                            audioFile.Dispose();
                            File.Delete(tempFile2);
                        };
                    }
                }
                else if (cmbTTSEngine.SelectedIndex == 3) // ElevenLabs
                {
                    if (string.IsNullOrEmpty(elevenLabsApiKey))
                    {
                        MessageBox.Show("Please set ElevenLabs API key first");
                        return;
                    }
                    
                    string tempFile3 = Path.GetTempFileName() + ".wav";
                    bool success3 = await CallElevenLabs(testText, tempFile3);
                    
                    if (success3 && File.Exists(tempFile3))
                    {
                        waveOut = new WaveOutEvent();
                        var audioFile = new AudioFileReader(tempFile3);
                        waveOut.Init(audioFile);
                        waveOut.Play();
                        waveOut.PlaybackStopped += (s, args) =>
                        {
                            audioFile.Dispose();
                            File.Delete(tempFile3);
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error testing voice: {ex.Message}");
            }
        }
        private async Task<bool> CallAWSPolly(string text, string outputFile)
        {
            try
            {
                if (pollyClient == null)
                {
                    LogMessage("AWS Polly client not initialized");
                    return false;
                }
                
                // Check if text contains SSML tags
                bool useSSML = ContainsSSMLTags(text);
                
                var request = new SynthesizeSpeechRequest
                {
                    OutputFormat = OutputFormat.Pcm,
                    VoiceId = currentAwsVoice,
                    Engine = useAwsNeuralEngine ? Engine.Neural : Engine.Standard,
                    SampleRate = "16000",
                    TextType = useSSML ? TextType.Ssml : TextType.Text
                };
                
                if (useSSML)
                {
                    // Convert to AWS-compatible SSML
                    string ssmlText = ConvertToAWSSSML(text);
                    request.Text = ssmlText;
                    LogMessage("Using SSML for AWS Polly test");
                }
                else
                {
                    request.Text = text;
                }
                
                LogMessage($"Calling AWS Polly with voice: {currentAwsVoice} ({(useAwsNeuralEngine ? "Neural" : "Standard")})");
                
                var response = await pollyClient.SynthesizeSpeechAsync(request);
                
                if (response.HttpStatusCode == System.Net.HttpStatusCode.OK)
                {
                    using (var responseStream = response.AudioStream)
                    using (var fs = new FileStream(outputFile, FileMode.Create))
                    using (var writer = new BinaryWriter(fs))
                    {
                        // Read PCM data
                        byte[] buffer = new byte[8192];
                        var pcmData = new List<byte>();
                        int bytesRead;
                        
                        while ((bytesRead = await responseStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            pcmData.AddRange(buffer.Take(bytesRead));
                        }
                        
                        // Write WAV header for 24000 Hz, 16-bit mono PCM
                       // Write WAV header - match the sample rate we requested
                        int sampleRate = 16000;  // Must match the SampleRate in the request
                        short bitsPerSample = 16;
                        short channels = 1;
                        int byteRate = sampleRate * channels * (bitsPerSample / 8);
                        short blockAlign = (short)(channels * (bitsPerSample / 8));
                        int dataSize = pcmData.Count;

                        writer.Write(Encoding.UTF8.GetBytes("RIFF"));
                        writer.Write(dataSize + 36);
                        writer.Write(Encoding.UTF8.GetBytes("WAVE"));
                        writer.Write(Encoding.UTF8.GetBytes("fmt "));
                        writer.Write(16);
                        writer.Write((short)1);  // PCM format
                        writer.Write(channels);
                        writer.Write(sampleRate);
                        writer.Write(byteRate);
                        writer.Write(blockAlign);
                        writer.Write(bitsPerSample);
                        writer.Write(Encoding.UTF8.GetBytes("data"));
                        writer.Write(dataSize);
                        writer.Write(pcmData.ToArray());
                    }
                    
                    LogMessage($"Successfully saved audio to {outputFile}");
                    return true;
                }
                
                LogMessage($"AWS Polly returned status: {response.HttpStatusCode}");
                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"AWS Polly Error: {ex.Message}");
                return false;
            }
        }

        private async Task<bool> CallAWSPollyWithSettings(string text, string outputFile, 
            string voiceToUse, bool useNeural, double rateValue, double volumeValue)
        {
            try
            {
                if (pollyClient == null)
                {
                    throw new Exception("AWS Polly client not initialized");
                }
                
                bool useSSML = ContainsSSMLTags(text);
                
                var request = new SynthesizeSpeechRequest
                {
                    OutputFormat = OutputFormat.Pcm,
                    VoiceId = voiceToUse,
                    Engine = useNeural ? Engine.Neural : Engine.Standard,
                    SampleRate = "16000",  // MUST be string "16000"
                    TextType = useSSML ? TextType.Ssml : TextType.Text
                };
                
                if (useSSML)
                {
                    string ssmlText = ConvertToAWSSSML(text);
                    
                    // Add prosody adjustments if needed
                    if (Math.Abs(rateValue) > 0.1 || Math.Abs(volumeValue - 100) > 1)
                    {
                        ssmlText = AddAWSProsody(ssmlText, rateValue, volumeValue);
                    }
                    
                    request.Text = ssmlText;
                    Dispatcher.Invoke(() => LogMessage("Using SSML for AWS Polly"));
                }
                else
                {
                    // Wrap plain text with prosody if needed
                    if (Math.Abs(rateValue) > 0.1 || Math.Abs(volumeValue - 100) > 1)
                    {
                        string prosodyText = WrapAWSProsody(text, rateValue, volumeValue);
                        request.Text = prosodyText;
                        request.TextType = TextType.Ssml;
                    }
                    else
                    {
                        request.Text = text;
                    }
                }
                
                Dispatcher.Invoke(() => LogMessage($"Calling AWS Polly with voice: {voiceToUse} ({(useNeural ? "Neural" : "Standard")})"));
                
                var response = await pollyClient.SynthesizeSpeechAsync(request);
                
                if (response.HttpStatusCode == System.Net.HttpStatusCode.OK)
                {
                    using (var responseStream = response.AudioStream)
                    using (var fs = new FileStream(outputFile, FileMode.Create))
                    using (var writer = new BinaryWriter(fs))
                    {
                        byte[] buffer = new byte[8192];
                        var pcmData = new List<byte>();
                        int bytesRead;
                        
                        while ((bytesRead = await responseStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            pcmData.AddRange(buffer.Take(bytesRead));
                        }
                        
                        // Write WAV header for 16000 Hz, 16-bit mono PCM
                        int sampleRate = 16000;  // Match the request sample rate
                        short bitsPerSample = 16;
                        short channels = 1;
                        int byteRate = sampleRate * channels * (bitsPerSample / 8);
                        short blockAlign = (short)(channels * (bitsPerSample / 8));
                        int dataSize = pcmData.Count;
                        
                        writer.Write(Encoding.UTF8.GetBytes("RIFF"));
                        writer.Write(dataSize + 36);
                        writer.Write(Encoding.UTF8.GetBytes("WAVE"));
                        writer.Write(Encoding.UTF8.GetBytes("fmt "));
                        writer.Write(16);
                        writer.Write((short)1);
                        writer.Write(channels);
                        writer.Write(sampleRate);
                        writer.Write(byteRate);
                        writer.Write(blockAlign);
                        writer.Write(bitsPerSample);
                        writer.Write(Encoding.UTF8.GetBytes("data"));
                        writer.Write(dataSize);
                        writer.Write(pcmData.ToArray());
                    }
                    
                    Dispatcher.Invoke(() => LogMessage($"Successfully saved audio to {outputFile}"));
                    return true;
                }
                
                Dispatcher.Invoke(() => LogMessage($"AWS Polly returned status: {response.HttpStatusCode}"));
                return false;
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => LogMessage($"AWS Polly Error: {ex.Message}"));
                return false;
            }
        }
        
        private void StopTest_Click(object sender, RoutedEventArgs e)
        {
            if (cmbTTSEngine.SelectedIndex == 0)
            {
                speechSynthesizer.SpeakAsyncCancelAll();
            }
            
            if (waveOut != null && waveOut.PlaybackState == PlaybackState.Playing)
            {
                waveOut.Stop();
            }
            
            isPlaying = false;
        }

        private async void TestSelection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string testText;
                var selection = rtbTextContent.Selection;
                
                if (!selection.IsEmpty)
                {
                    testText = selection.Text;
                    LogMessage($"Testing selected text: {testText.Substring(0, Math.Min(50, testText.Length))}...");
                }
                else
                {
                    // Get text around cursor (50 chars before and after)
                    var caretPos = rtbTextContent.CaretPosition;
                    var start = caretPos.GetPositionAtOffset(-50) ?? rtbTextContent.Document.ContentStart;
                    var end = caretPos.GetPositionAtOffset(50) ?? rtbTextContent.Document.ContentEnd;
                    testText = new TextRange(start, end).Text;
                    LogMessage("Testing text around cursor position");
                }
                
                if (string.IsNullOrWhiteSpace(testText))
                {
                    MessageBox.Show("No text selected or around cursor to test.");
                    return;
                }
                
                // Stop any current playback
                StopTestSelection_Click(sender, e);
                
                // Enable stop button
                btnStopTestSelection.IsEnabled = true;
                btnTestSelection.IsEnabled = false;
                
                // Remove control tags (split and voice tags) for testing
                string cleanedText = testText;
                
                // Remove <split> tags
                cleanedText = Regex.Replace(cleanedText, @"<split>", "", RegexOptions.IgnoreCase);

                // Remove <service> tags
                cleanedText = Regex.Replace(cleanedText, @"<service=\d+>", "", RegexOptions.IgnoreCase);

                // Extract voice index if present and remove voice tags
                int testVoiceIndex = -1;
                var voiceMatch = Regex.Match(cleanedText, @"<voice=(\d+)>");
                if (voiceMatch.Success)
                {
                    testVoiceIndex = int.Parse(voiceMatch.Groups[1].Value) - 1;
                    cleanedText = Regex.Replace(cleanedText, @"<voice=\d+>", "", RegexOptions.IgnoreCase);
                }
                
                // Trim the cleaned text
                cleanedText = cleanedText.Trim();
                
                if (string.IsNullOrWhiteSpace(cleanedText))
                {
                    MessageBox.Show("No content to test after removing control tags.");
                    btnStopTestSelection.IsEnabled = false;
                    btnTestSelection.IsEnabled = true;
                    return;
                }
                
                if (cmbTTSEngine.SelectedIndex == 0) // Windows SAPI
                {
                    // Save current voice selection
                    string originalVoice = cmbVoices.SelectedItem?.ToString();
                    
                    // If a voice index was found in the selection, use that voice
                    if (testVoiceIndex >= 0)
                    {
                        var voices = speechSynthesizer.GetInstalledVoices();
                        if (testVoiceIndex < voices.Count)
                        {
                            speechSynthesizer.SelectVoice(voices[testVoiceIndex].VoiceInfo.Name);
                            LogMessage($"Testing with voice: {voices[testVoiceIndex].VoiceInfo.Name}");
                        }
                    }
                    
                    speechSynthesizer.Rate = (int)sliderRate.Value;
                    speechSynthesizer.Volume = (int)sliderVolume.Value;
                    
                    // Create a completion handler to restore voice
                    EventHandler<SpeakCompletedEventArgs> completedHandler = null;
                    completedHandler = (s, args) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            btnStopTestSelection.IsEnabled = false;
                            btnTestSelection.IsEnabled = true;
                            
                            // Restore original voice
                            if (!string.IsNullOrEmpty(originalVoice))
                            {
                                speechSynthesizer.SelectVoice(originalVoice);
                            }
                            
                            speechSynthesizer.SpeakCompleted -= completedHandler;
                        });
                    };
                    
                    speechSynthesizer.SpeakCompleted += completedHandler;
                    
                    if (ContainsSSMLTags(cleanedText))
                    {
                        string ssmlText = WrapInSSML(cleanedText);
                        LogMessage("Testing with SSML (SAPI)");
                        speechSynthesizer.SpeakSsmlAsync(ssmlText);
                    }
                    else
                    {
                        speechSynthesizer.SpeakAsync(cleanedText);
                    }
                    
                    isPlaying = true;
                }
                else if (cmbTTSEngine.SelectedIndex == 1) // Google TTS
                {
                    if (string.IsNullOrEmpty(googleApiKey))
                    {
                        MessageBox.Show("Please set Google API key first");
                        btnStopTestSelection.IsEnabled = false;
                        btnTestSelection.IsEnabled = true;
                        return;
                    }
                    
                    // Save current voice selection
                    string originalGoogleVoice = currentGoogleVoice;
                    
                    // If a voice index was found, use that voice
                    if (testVoiceIndex >= 0)
                    {
                        var voicesList = googleVoices.Values.ToList();
                        if (testVoiceIndex < voicesList.Count)
                        {
                            currentGoogleVoice = voicesList[testVoiceIndex];
                            LogMessage($"Testing with Google voice: {currentGoogleVoice}");
                        }
                    }
                    
                    string tempFile = Path.GetTempFileName() + ".wav";
                    bool success = await CallGoogleTTS(cleanedText, tempFile);
                    
                    // Restore original voice
                    currentGoogleVoice = originalGoogleVoice;
                    
                    if (success && File.Exists(tempFile))
                    {
                        waveOut = new WaveOutEvent();
                        var audioFile = new AudioFileReader(tempFile);
                        waveOut.Init(audioFile);
                        waveOut.Play();
                        waveOut.PlaybackStopped += (s, args) =>
                        {
                            audioFile.Dispose();
                            File.Delete(tempFile);
                            Dispatcher.Invoke(() =>
                            {
                                btnStopTestSelection.IsEnabled = false;
                                btnTestSelection.IsEnabled = true;
                            });
                        };
                        isPlaying = true;
                    }
                    else
                    {
                        btnStopTestSelection.IsEnabled = false;
                        btnTestSelection.IsEnabled = true;
                    }
                }
                else if (cmbTTSEngine.SelectedIndex == 2) // AWS Polly
                {
                    if (pollyClient == null)
                    {
                        MessageBox.Show("Please set AWS credentials first");
                        btnStopTestSelection.IsEnabled = false;
                        btnTestSelection.IsEnabled = true;
                        return;
                    }
                    
                    // Save current voice selection
                    string originalAwsVoice = currentAwsVoice;
                    bool originalUseNeural = useAwsNeuralEngine;
                    
                    // If a voice index was found, use that voice
                    if (testVoiceIndex >= 0)
                    {
                        var voicesList = awsVoices.Values.ToList();
                        if (testVoiceIndex < voicesList.Count)
                        {
                            currentAwsVoice = voicesList[testVoiceIndex].VoiceId;
                            useAwsNeuralEngine = voicesList[testVoiceIndex].Engine == "neural";
                            LogMessage($"Testing with AWS voice: {currentAwsVoice} ({(useAwsNeuralEngine ? "Neural" : "Standard")})");
                        }
                    }
                    
                    string tempFile = Path.GetTempFileName() + ".wav";
                    bool success = await CallAWSPolly(cleanedText, tempFile);
                    
                    // Restore original voice
                    currentAwsVoice = originalAwsVoice;
                    useAwsNeuralEngine = originalUseNeural;
                    
                    if (success && File.Exists(tempFile))
                    {
                        waveOut = new WaveOutEvent();
                        var audioFile = new AudioFileReader(tempFile);
                        waveOut.Init(audioFile);
                        waveOut.Play();
                        waveOut.PlaybackStopped += (s, args) =>
                        {
                            audioFile.Dispose();
                            File.Delete(tempFile);
                            Dispatcher.Invoke(() =>
                            {
                                btnStopTestSelection.IsEnabled = false;
                                btnTestSelection.IsEnabled = true;
                            });
                        };
                        isPlaying = true;
                    }
                    else
                    {
                        btnStopTestSelection.IsEnabled = false;
                        btnTestSelection.IsEnabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error testing selection: {ex.Message}");
                btnStopTestSelection.IsEnabled = false;
                btnTestSelection.IsEnabled = true;
            }
        }
        
        private async void Convert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
                if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;
                
                string outputPath = folderDialog.SelectedPath;
                lastOutputFolder = outputPath;
                
                // Get text on UI thread
                string text = new TextRange(rtbTextContent.Document.ContentStart, 
                    rtbTextContent.Document.ContentEnd).Text;
                
                // Get current settings on UI thread
                int engineIndex = cmbTTSEngine.SelectedIndex;
                int outputFormatIndex = cmbOutputFormat.SelectedIndex;
                double rateValue = sliderRate.Value;
                double volumeValue = sliderVolume.Value;
                
                // Show progress
                progressBar.Visibility = Visibility.Visible;
                txtStatus.Text = "Converting...";
        // Run conversion in background
                await Task.Run(async () =>
                {
                    // Clear previous file list
                    createdAudioFiles.Clear();
                    
                    // Process splits and voice changes
                    var segments = ProcessTextSegments(text);
                    
                    Dispatcher.Invoke(() => LogMessage($"Processing {segments.Count} segments"));
                    
                    // Group segments by SplitIndex
                    var segmentGroups = segments.GroupBy(s => s.SplitIndex).OrderBy(g => g.Key).ToList();
                    
                    for (int groupIndex = 0; groupIndex < segmentGroups.Count; groupIndex++)
                    {
                        var group = segmentGroups[groupIndex];
                        var groupSegments = group.ToList();
                        
                        Dispatcher.Invoke(() => 
                        {
                            progressBar.Value = (groupIndex * 100) / segmentGroups.Count;
                        });
                        
                        string baseFileName = $"output_{groupIndex + 1:D3}";
                        var tempFiles = new List<string>();
                        
                        // Process each sub-segment
                        for (int subIndex = 0; subIndex < groupSegments.Count; subIndex++)
                        {
                            var segment = groupSegments[subIndex];
                            
                            string subFileName = groupSegments.Count > 1 
                                ? $"{baseFileName}{(char)('a' + subIndex)}" 
                                : baseFileName;
                            
                            string outputFile = Path.Combine(outputPath, subFileName);
                            
                            // Determine which service to use
                            int serviceToUse = segment.ServiceIndex >= 0 ? segment.ServiceIndex : engineIndex;
                            
                            Dispatcher.Invoke(() => 
                                LogMessage($"Segment {groupIndex + 1}.{subIndex + 1}: File={subFileName}, Service={serviceToUse + 1}, Voice={segment.VoiceIndex}, Length={segment.Text.Length} chars"));
                            
                            try
                            {
                                if (serviceToUse == 0) // Windows SAPI
                                {
                                    await ConvertWithSAPIBackground(segment, outputFile, rateValue, volumeValue, outputFormatIndex);
                                }
                                else if (serviceToUse == 1) // Google Cloud TTS
                                {
                                    await ConvertWithGoogleTTSBackground(segment, outputFile, rateValue, volumeValue, outputFormatIndex);
                                }
                                else if (serviceToUse == 2) // AWS Polly
                                {
                                    await ConvertWithAWSPollyBackground(segment, outputFile, rateValue, volumeValue, outputFormatIndex);
                                }
                                else if (serviceToUse == 3) // ElevenLabs
                                {
                                    await ConvertWithElevenLabsBackground(segment, outputFile, rateValue, volumeValue, outputFormatIndex);
                                }
                                
                                string extension = outputFormatIndex == 0 ? ".wav" : ".mp3";
                                string fullPath = outputFile + extension;
                                
                                if (File.Exists(fullPath))
                                {
                                    tempFiles.Add(fullPath);
                                    Dispatcher.Invoke(() => LogMessage($"Created: {fullPath}"));
                                }
                            }
                            catch (Exception ex)
                            {
                                Dispatcher.Invoke(() => LogMessage($"Error converting segment: {ex.Message}"));
                            }
                        }
                        
                        // Merge sub-files if there are multiple
                        string finalFile = Path.Combine(outputPath, baseFileName + (outputFormatIndex == 0 ? ".wav" : ".mp3"));
                        
                        if (tempFiles.Count > 1)
                        {
                            Dispatcher.Invoke(() => LogMessage($"Merging {tempFiles.Count} sub-files into {baseFileName}..."));
                            
                            try
                            {
                                await MergeAudioFiles(tempFiles, finalFile);
                                
                                // Delete temp files after merging
                                foreach (var tempFile in tempFiles)
                                {
                                    File.Delete(tempFile);
                                }
                                
                                Dispatcher.Invoke(() => LogMessage($"Merged into: {finalFile}"));
                            }
                            catch (Exception ex)
                            {
                                Dispatcher.Invoke(() => LogMessage($"Error merging files: {ex.Message}"));
                            }
                        }
                        else if (tempFiles.Count == 1)
                        {
                            // Only one file, just rename it if needed
                            if (tempFiles[0] != finalFile)
                            {
                                File.Move(tempFiles[0], finalFile);
                            }
                        }
                        
                        if (File.Exists(finalFile))
                        {
                            createdAudioFiles.Add(finalFile);
                        }
                    }
                    
                    Dispatcher.Invoke(() =>
                    {
                        progressBar.Visibility = Visibility.Collapsed;
                        txtStatus.Text = "Conversion complete";
                        LogMessage($"All {segmentGroups.Count} files converted successfully!");
                        
                        // Enable playback controls
                        UpdatePlaybackControls();
                    });
                });
            }
            catch (Exception ex)
            {
                LogMessage($"Error during conversion: {ex.Message}");
                progressBar.Visibility = Visibility.Collapsed;
                txtStatus.Text = "Error";
                MessageBox.Show($"Conversion error: {ex.Message}");
            }
        }   
        private async Task MergeAudioFiles(List<string> files, string outputFile)
        {
            await Task.Run(() =>
            {
                // Determine the output format
                WaveFormat outputFormat = null;
                
                // Get format from first file
                using (var reader = new AudioFileReader(files[0]))
                {
                    outputFormat = reader.WaveFormat;
                }
                
                using (var writer = new WaveFileWriter(outputFile, outputFormat))
                {
                    foreach (var file in files)
                    {
                        using (var reader = new AudioFileReader(file))
                        {
                            // Resample if formats don't match
                            if (reader.WaveFormat.SampleRate != outputFormat.SampleRate ||
                                reader.WaveFormat.Channels != outputFormat.Channels)
                            {
                                using (var resampler = new MediaFoundationResampler(reader, outputFormat))
                                {
                                    resampler.ResamplerQuality = 60;
                                    byte[] buffer = new byte[outputFormat.AverageBytesPerSecond];
                                    int bytesRead;
                                    while ((bytesRead = resampler.Read(buffer, 0, buffer.Length)) > 0)
                                    {
                                        writer.Write(buffer, 0, bytesRead);
                                    }
                                }
                            }
                            else
                            {
                                // Same format, just copy
                                reader.CopyTo(writer);
                            }
                        }
                    }
                }
            });
        }     
        private void StopTestSelection_Click(object sender, RoutedEventArgs e)
        {
            if (cmbTTSEngine.SelectedIndex == 0)
            {
                speechSynthesizer.SpeakAsyncCancelAll();
            }
            
            if (waveOut != null && waveOut.PlaybackState == PlaybackState.Playing)
            {
                waveOut.Stop();
            }
            
            isPlaying = false;
            btnStopTestSelection.IsEnabled = false;
            btnTestSelection.IsEnabled = true;
        }

        private string ConvertToAWSSSML(string text)
        {
            // If already wrapped in speak tags, extract content
            if (text.TrimStart().StartsWith("<speak", StringComparison.OrdinalIgnoreCase))
            {
                var match = Regex.Match(text, @"<speak[^>]*>(.*?)</speak>", 
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (match.Success)
                {
                    text = match.Groups[1].Value;
                }
            }
            
            // AWS Polly SSML compatibility adjustments
            // Fix pitch values - AWS accepts percentage or Hz
            text = Regex.Replace(text, @"pitch=""high""", @"pitch=""+20%""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""x-high""", @"pitch=""+40%""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""low""", @"pitch=""-20%""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""x-low""", @"pitch=""-40%""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""medium""", @"pitch=""+0%""", RegexOptions.IgnoreCase);
            
            // Convert semitones to percentage for AWS
            text = Regex.Replace(text, @"pitch=""([+-]?\d+)st""", match =>
            {
                int semitones = int.Parse(match.Groups[1].Value);
                int percentage = semitones * 8; // Approximate conversion
                return $"pitch=\"{(percentage >= 0 ? "+" : "")}{percentage}%\"";
            }, RegexOptions.IgnoreCase);
            
            return $"<speak>{text}</speak>";
        }

        private string WrapAWSProsody(string text, double rateValue, double volumeValue)
        {
            var prosodyAttrs = new List<string>();
            
            // Convert rate (-10 to 10) to percentage
            if (Math.Abs(rateValue) > 0.1)
            {
                int ratePercent = (int)(100 + (rateValue * 10));
                prosodyAttrs.Add($"rate=\"{ratePercent}%\"");
            }
            
            // Convert volume (0 to 100) to dB or descriptor
            if (Math.Abs(volumeValue - 100) > 1)
            {
                string volumeStr;
                if (volumeValue >= 80)
                    volumeStr = "x-loud";
                else if (volumeValue >= 60)
                    volumeStr = "loud";
                else if (volumeValue >= 40)
                    volumeStr = "medium";
                else if (volumeValue >= 20)
                    volumeStr = "soft";
                else
                    volumeStr = "x-soft";
                
                prosodyAttrs.Add($"volume=\"{volumeStr}\"");
            }
            
            if (prosodyAttrs.Count > 0)
            {
                string attrs = string.Join(" ", prosodyAttrs);
                return $"<speak><prosody {attrs}>{text}</prosody></speak>";
            }
            
            return $"<speak>{text}</speak>";
        }

        private string AddAWSProsody(string ssmlText, double rateValue, double volumeValue)
        {
            // Extract content from speak tags
            var match = Regex.Match(ssmlText, @"<speak>(.*?)</speak>", 
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            
            if (!match.Success)
                return ssmlText;
            
            string content = match.Groups[1].Value;
            
            var prosodyAttrs = new List<string>();
            
            if (Math.Abs(rateValue) > 0.1)
            {
                int ratePercent = (int)(100 + (rateValue * 10));
                prosodyAttrs.Add($"rate=\"{ratePercent}%\"");
            }
            
            if (Math.Abs(volumeValue - 100) > 1)
            {
                string volumeStr;
                if (volumeValue >= 80)
                    volumeStr = "x-loud";
                else if (volumeValue >= 60)
                    volumeStr = "loud";
                else if (volumeValue >= 40)
                    volumeStr = "medium";
                else if (volumeValue >= 20)
                    volumeStr = "soft";
                else
                    volumeStr = "x-soft";
                
                prosodyAttrs.Add($"volume=\"{volumeStr}\"");
            }
            
            if (prosodyAttrs.Count > 0)
            {
                string attrs = string.Join(" ", prosodyAttrs);
                return $"<speak><prosody {attrs}>{content}</prosody></speak>";
            }
            
            return ssmlText;
        }

        
        private void HighlightSSML_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var document = rtbTextContent.Document;
                var textRange = new TextRange(document.ContentStart, document.ContentEnd);
                string text = textRange.Text;
                
                // Clear existing formatting
                textRange.ClearAllProperties();
                
                // Pattern to match SSML tags
                string pattern = @"</?(?:emphasis|break|prosody|say-as|phoneme|sub|audio|p|s|speak|voice|mark|desc|lexicon|metadata|meta)[^>]*>";
                var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                
                int highlightCount = 0;
                foreach (Match match in matches)
                {
                    var start = GetTextPointerAtOffset(document.ContentStart, match.Index);
                    var end = GetTextPointerAtOffset(document.ContentStart, match.Index + match.Length);
                    
                    if (start != null && end != null)
                    {
                        var range = new TextRange(start, end);
                        range.ApplyPropertyValue(TextElement.BackgroundProperty, Brushes.Yellow);
                        range.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.DarkBlue);
                        range.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                        highlightCount++;
                    }
                }
                
                // Also highlight <split> and <voice> tags
                pattern = @"<(?:split|voice=\d+|service=\d+)>";
                matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);

                foreach (Match match in matches)
                {
                    var start = GetTextPointerAtOffset(document.ContentStart, match.Index);
                    var end = GetTextPointerAtOffset(document.ContentStart, match.Index + match.Length);
                    
                    if (start != null && end != null)
                    {
                        var range = new TextRange(start, end);
                        range.ApplyPropertyValue(TextElement.BackgroundProperty, Brushes.LightGreen);
                        range.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.DarkGreen);
                        range.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                        highlightCount++;
                    }
                }
                
                LogMessage($"Highlighted {highlightCount} tags (Yellow=SSML, Green=Control)");
            }
            catch (Exception ex)
            {
                LogMessage($"Error highlighting SSML: {ex.Message}");
            }
        }
        
        private void ValidateSSML_Click(object sender, RoutedEventArgs e)
        {
            var textRange = new TextRange(rtbTextContent.Document.ContentStart, rtbTextContent.Document.ContentEnd);
            string text = textRange.Text;
            
            var validationResult = ValidateSSMLSyntax(text);
            
            // Clear existing formatting
            textRange.ClearAllProperties();
            
            // Display validation results
            if (validationResult.Errors.Count == 0)
            {
                LogMessage("✓ SSML validation passed! No errors found.");
                MessageBox.Show("SSML validation successful!\n\nAll tags are properly formed and closed.", 
                    "Validation Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                // Highlight errors
                foreach (var error in validationResult.Errors)
                {
                    HighlightError(error.Position, error.Length);
                }
                
                // Show detailed error report
                var errorReport = new StringBuilder();
                errorReport.AppendLine($"Found {validationResult.Errors.Count} error(s):\n");
                
                int errorNum = 1;
                foreach (var error in validationResult.Errors)
                {
                    errorReport.AppendLine($"{errorNum}. {error.Message}");
                    if (!string.IsNullOrEmpty(error.Context))
                    {
                        errorReport.AppendLine($"   Context: {error.Context}");
                    }
                    errorReport.AppendLine($"   Position: Character {error.Position}");
                    errorReport.AppendLine();
                    errorNum++;
                }
                
                LogMessage($"✗ SSML validation failed with {validationResult.Errors.Count} error(s)");
                
                // Create a detailed error window
                var errorWindow = new Window
                {
                    Title = "SSML Validation Errors",
                    Width = 600,
                    Height = 400,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };
                
                var grid = new Grid();
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                
                var errorTextBox = new TextBox
                {
                    Text = errorReport.ToString(),
                    IsReadOnly = true,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    FontFamily = new FontFamily("Consolas"),
                    Margin = new Thickness(10)
                };
                
                var closeButton = new Button
                {
                    Content = "Close",
                    Width = 100,
                    Height = 30,
                    Margin = new Thickness(10),
                    HorizontalAlignment = HorizontalAlignment.Right
                };
                
                closeButton.Click += (s, args) => errorWindow.Close();
                
                Grid.SetRow(errorTextBox, 0);
                Grid.SetRow(closeButton, 1);
                
                grid.Children.Add(errorTextBox);
                grid.Children.Add(closeButton);
                
                errorWindow.Content = grid;
                errorWindow.ShowDialog();
            }
            
            // Also show warnings if any
            if (validationResult.Warnings.Count > 0)
            {
                var warningText = new StringBuilder();
                warningText.AppendLine("\nWarnings:");
                foreach (var warning in validationResult.Warnings)
                {
                    warningText.AppendLine($"⚠ {warning}");
                }
                LogMessage(warningText.ToString());
            }
        }
        

        
        // Part 2: Helper Methods
        private List<TextSegment> ProcessTextSegments(string text)
        {
            var segments = new List<TextSegment>();
            
            // Split by <split> tags first
            var splitPattern = @"<split>";
            var majorParts = Regex.Split(text, splitPattern, RegexOptions.IgnoreCase);
            
            for (int splitIndex = 0; splitIndex < majorParts.Length; splitIndex++)
            {
                var part = majorParts[splitIndex];
                if (string.IsNullOrWhiteSpace(part)) continue;
                
                // Within each split section, process service/voice changes
                var subSegments = ProcessServiceAndVoiceChanges(part, splitIndex);
                segments.AddRange(subSegments);
            }
            
            if (segments.Count == 0)
            {
                segments.Add(new TextSegment 
                { 
                    Text = text, 
                    VoiceIndex = 0,
                    ServiceIndex = -1,
                    SplitIndex = 0,
                    SubIndex = 0
                });
            }
            
            return segments;
        }

        private List<TextSegment> ProcessServiceAndVoiceChanges(string text, int splitIndex)
        {
            var segments = new List<TextSegment>();
            int currentVoiceIndex = 0;
            int currentServiceIndex = -1;
            int subIndex = 0;
            
            // Find all service and voice tags
            var pattern = @"<(voice|service)=(\d+)>";
            var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
            
            if (matches.Count == 0)
            {
                // No service/voice changes, just one segment
                var cleanText = text.Trim();
                if (!string.IsNullOrEmpty(cleanText))
                {
                    segments.Add(new TextSegment 
                    { 
                        Text = cleanText, 
                        VoiceIndex = currentVoiceIndex,
                        ServiceIndex = currentServiceIndex,
                        SplitIndex = splitIndex,
                        SubIndex = subIndex
                    });
                }
                return segments;
            }
            
            int lastIndex = 0;
            
            foreach (Match match in matches)
            {
                // Add text before the tag
                if (match.Index > lastIndex)
                {
                    var beforeText = text.Substring(lastIndex, match.Index - lastIndex).Trim();
                    if (!string.IsNullOrEmpty(beforeText))
                    {
                        segments.Add(new TextSegment 
                        { 
                            Text = beforeText, 
                            VoiceIndex = currentVoiceIndex,
                            ServiceIndex = currentServiceIndex,
                            SplitIndex = splitIndex,
                            SubIndex = subIndex
                        });
                        subIndex++;
                    }
                }
                
                // Process the tag
                string tagType = match.Groups[1].Value.ToLower();
                int tagValue = int.Parse(match.Groups[2].Value);
                
                if (tagType == "voice")
                {
                    currentVoiceIndex = tagValue - 1;
                }
                else if (tagType == "service")
                {
                    currentServiceIndex = tagValue - 1;
                }
                
                lastIndex = match.Index + match.Length;
            }
            
            // Add remaining text after last tag
            if (lastIndex < text.Length)
            {
                var remainingText = text.Substring(lastIndex).Trim();
                if (!string.IsNullOrEmpty(remainingText))
                {
                    segments.Add(new TextSegment 
                    { 
                        Text = remainingText, 
                        VoiceIndex = currentVoiceIndex,
                        ServiceIndex = currentServiceIndex,
                        SplitIndex = splitIndex,
                        SubIndex = subIndex
                    });
                }
            }
            
            return segments;
        }
        
        private async Task ConvertWithSAPIBackground(TextSegment segment, string outputFile, 
            double rateValue, double volumeValue, int outputFormatIndex)
        {
            await Task.Run(() =>
            {
                using (var synth = new SpeechSynthesizer())
                {
                    var voices = synth.GetInstalledVoices();
                    
                    // Select the appropriate voice based on the segment's voice index
                    if (segment.VoiceIndex >= 0 && segment.VoiceIndex < voices.Count)
                    {
                        synth.SelectVoice(voices[segment.VoiceIndex].VoiceInfo.Name);
                        Dispatcher.Invoke(() => 
                            LogMessage($"SAPI: Using voice {voices[segment.VoiceIndex].VoiceInfo.Name} (index {segment.VoiceIndex})"));
                    }
                    else if (voices.Count > 0)
                    {
                        synth.SelectVoice(voices[0].VoiceInfo.Name);
                        Dispatcher.Invoke(() => 
                            LogMessage($"SAPI: Using default voice {voices[0].VoiceInfo.Name}"));
                    }
                    
                    synth.Rate = (int)rateValue;
                    synth.Volume = (int)volumeValue;
                    
                    string extension = outputFormatIndex == 0 ? ".wav" : ".mp3";
                    string wavFile = outputFile + ".wav";
                    
                    synth.SetOutputToWaveFile(wavFile);
                    
                    // Check if text contains SSML tags
                    if (ContainsSSMLTags(segment.Text))
                    {
                        // Speak as SSML
                        string ssmlText = WrapInSSML(segment.Text);
                        Dispatcher.Invoke(() => LogMessage("Using SSML for SAPI"));
                        synth.SpeakSsml(ssmlText);
                    }
                    else
                    {
                        // Speak as plain text
                        synth.Speak(segment.Text);
                    }
                    
                    synth.SetOutputToDefaultAudioDevice();
                    
                    if (extension == ".mp3")
                    {
                        ConvertWavToMp3(wavFile, outputFile + ".mp3");
                        File.Delete(wavFile);
                    }
                }
            });
        }
        private async Task ConvertWithAWSPollyBackground(TextSegment segment, string outputFile,
            double rateValue, double volumeValue, int outputFormatIndex)
        {
            if (pollyClient == null)
            {
                throw new Exception("AWS Polly client not initialized");
            }
            
            var voicesList = awsVoices.Values.ToList();
            
            // Default to first voice
            string voiceToUse = voicesList.Count > 0 ? voicesList[0].VoiceId : "Joanna";
            bool useNeural = voicesList.Count > 0 ? voicesList[0].Engine == "neural" : false;
            
            if (segment.VoiceIndex >= 0 && segment.VoiceIndex < voicesList.Count)
            {
                voiceToUse = voicesList[segment.VoiceIndex].VoiceId;
                useNeural = voicesList[segment.VoiceIndex].Engine == "neural";
                Dispatcher.Invoke(() => 
                    LogMessage($"AWS Polly: Using voice {voiceToUse} ({(useNeural ? "Neural" : "Standard")}) (index {segment.VoiceIndex})"));
            }
            else
            {
                Dispatcher.Invoke(() => 
                    LogMessage($"AWS Polly: Using default voice {voiceToUse}"));
            }
            
            string extension = outputFormatIndex == 0 ? ".wav" : ".mp3";
            string wavFile = outputFile + ".wav";
            
            bool success = await CallAWSPollyWithSettings(segment.Text, wavFile, voiceToUse, useNeural, rateValue, volumeValue);
            
            if (!success)
            {
                throw new Exception("AWS Polly API call failed");
            }
            
            if (extension == ".mp3")
            {
                ConvertWavToMp3(wavFile, outputFile + ".mp3");
                File.Delete(wavFile);
            }
        }
        private async Task ConvertWithGoogleTTSBackground(TextSegment segment, string outputFile,
            double rateValue, double volumeValue, int outputFormatIndex)
        {
            if (string.IsNullOrEmpty(googleApiKey))
            {
                throw new Exception("Google API key not set");
            }
            
            // Get Google voices list
            var voicesList = googleVoices.Values.ToList();
            
            // Select the appropriate voice based on segment's voice index
            string voiceToUse = voicesList.Count > 0 ? voicesList[0] : "en-US-Wavenet-D"; // Default
            
            if (segment.VoiceIndex >= 0 && segment.VoiceIndex < voicesList.Count)
            {
                voiceToUse = voicesList[segment.VoiceIndex];
                Dispatcher.Invoke(() => 
                    LogMessage($"Google TTS: Using voice {voiceToUse} (index {segment.VoiceIndex})"));
            }
            else
            {
                Dispatcher.Invoke(() => 
                    LogMessage($"Google TTS: Using default voice {voiceToUse}"));
            }
            
            string extension = outputFormatIndex == 0 ? ".wav" : ".mp3";
            string wavFile = outputFile + ".wav";
            
            bool success = await CallGoogleTTSWithSettings(segment.Text, wavFile, voiceToUse, rateValue, volumeValue);
            
            if (!success)
            {
                throw new Exception("Google TTS API call failed");
            }
            
            if (extension == ".mp3")
            {
                ConvertWavToMp3(wavFile, outputFile + ".mp3");
                File.Delete(wavFile);
            }
        }
        
        private async Task<bool> CallGoogleTTS(string text, string outputFile)
        {
            try
            {
                string url = $"https://texttospeech.googleapis.com/v1/text:synthesize?key={googleApiKey}";
                
                // Determine gender based on voice
                string gender = "NEUTRAL";
                if (currentGoogleVoice.Contains("-C") || currentGoogleVoice.Contains("-E") || 
                    currentGoogleVoice.Contains("-F") || currentGoogleVoice.Contains("-H"))
                {
                    gender = "FEMALE";
                }
                else if (currentGoogleVoice.Contains("-A") || currentGoogleVoice.Contains("-B") || 
                         currentGoogleVoice.Contains("-D") || currentGoogleVoice.Contains("-I") || 
                         currentGoogleVoice.Contains("-J"))
                {
                    gender = "MALE";
                }
                
                // Check if text contains SSML tags
                bool useSSML = ContainsSSMLTags(text);
                object inputObject;
                
                if (useSSML)
                {
                    // Process as SSML
                    string ssmlText = ConvertToGoogleSSML(text);
                    inputObject = new { ssml = ssmlText };
                    LogMessage("Using SSML for Google TTS test");
                }
                else
                {
                    // Process as plain text
                    inputObject = new { text = text };
                }
                
                var requestBody = new
                {
                    input = inputObject,
                    voice = new
                    {
                        languageCode = "en-US",
                        name = currentGoogleVoice,
                        ssmlGender = gender
                    },
                    audioConfig = new
                    {
                        audioEncoding = "LINEAR16",
                        speakingRate = 1.0 + (sliderRate.Value / 10.0),
                        pitch = 0.0,
                        volumeGainDb = (sliderVolume.Value - 100) / 5.0
                    }
                };
                
                var json = JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                LogMessage($"Calling Google TTS with voice: {currentGoogleVoice}");
                
                var response = await httpClient.PostAsync(url, content);
                
                if (!response.IsSuccessStatusCode)
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    LogMessage($"Google TTS Error {response.StatusCode}: {errorContent}");
                    return false;
                }
                
                var responseJson = await response.Content.ReadAsStringAsync();
                var responseData = JsonSerializer.Deserialize<Dictionary<string, object>>(responseJson);
                
                if (responseData.ContainsKey("audioContent"))
                {
                    string audioContent = responseData["audioContent"].ToString();
                    byte[] audioBytes = Convert.FromBase64String(audioContent);
                    
                    // Write WAV header and data
                    using (var fs = new FileStream(outputFile, FileMode.Create))
                    using (var writer = new BinaryWriter(fs))
                    {
                        // WAV header for 24000 Hz, 16-bit mono
                        int sampleRate = 24000;
                        short bitsPerSample = 16;
                        short channels = 1;
                        int byteRate = sampleRate * channels * (bitsPerSample / 8);
                        short blockAlign = (short)(channels * (bitsPerSample / 8));
                        int dataSize = audioBytes.Length;
                        
                        writer.Write(Encoding.UTF8.GetBytes("RIFF"));
                        writer.Write(dataSize + 36);
                        writer.Write(Encoding.UTF8.GetBytes("WAVE"));
                        writer.Write(Encoding.UTF8.GetBytes("fmt "));
                        writer.Write(16);
                        writer.Write((short)1);
                        writer.Write(channels);
                        writer.Write(sampleRate);
                        writer.Write(byteRate);
                        writer.Write(blockAlign);
                        writer.Write(bitsPerSample);
                        writer.Write(Encoding.UTF8.GetBytes("data"));
                        writer.Write(dataSize);
                        writer.Write(audioBytes);
                    }
                    
                    LogMessage($"Successfully saved audio to {outputFile}");
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Google TTS Error: {ex.Message}");
                return false;
            }
        }
        
        private async Task<bool> CallGoogleTTSWithSettings(string text, string outputFile, 
            string voiceToUse, double rateValue, double volumeValue)
        {
            try
            {
                string url = $"https://texttospeech.googleapis.com/v1/text:synthesize?key={googleApiKey}";
                
                // Determine gender based on voice
                string gender = "NEUTRAL";
                if (voiceToUse.Contains("-C") || voiceToUse.Contains("-E") || 
                    voiceToUse.Contains("-F") || voiceToUse.Contains("-H"))
                {
                    gender = "FEMALE";
                }
                else if (voiceToUse.Contains("-A") || voiceToUse.Contains("-B") || 
                         voiceToUse.Contains("-D") || voiceToUse.Contains("-I") || 
                         voiceToUse.Contains("-J"))
                {
                    gender = "MALE";
                }
                
                // Check if text contains SSML tags
                bool useSSML = ContainsSSMLTags(text);
                object inputObject;
                
                if (useSSML)
                {
                    // Process as SSML
                    string ssmlText = ConvertToGoogleSSML(text);
                    inputObject = new { ssml = ssmlText };
                    Dispatcher.Invoke(() => LogMessage("Using SSML for Google TTS"));
                }
                else
                {
                    // Process as plain text
                    inputObject = new { text = text };
                }
                
                var requestBody = new
                {
                    input = inputObject,
                    voice = new
                    {
                        languageCode = "en-US",
                        name = voiceToUse,
                        ssmlGender = gender
                    },
                    audioConfig = new
                    {
                        audioEncoding = "LINEAR16",
                        speakingRate = 1.0 + (rateValue / 10.0),
                        pitch = 0.0,
                        volumeGainDb = (volumeValue - 100) / 5.0
                    }
                };
                
                var json = JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                Dispatcher.Invoke(() => LogMessage($"Calling Google TTS with voice: {voiceToUse}"));
                
                var response = await httpClient.PostAsync(url, content);
                
                if (!response.IsSuccessStatusCode)
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    Dispatcher.Invoke(() => LogMessage($"Google TTS Error {response.StatusCode}: {errorContent}"));
                    return false;
                }
                
                var responseJson = await response.Content.ReadAsStringAsync();
                var responseData = JsonSerializer.Deserialize<Dictionary<string, object>>(responseJson);
                
                if (responseData.ContainsKey("audioContent"))
                {
                    string audioContent = responseData["audioContent"].ToString();
                    byte[] audioBytes = Convert.FromBase64String(audioContent);
                    
                    // Write WAV header and data
                    using (var fs = new FileStream(outputFile, FileMode.Create))
                    using (var writer = new BinaryWriter(fs))
                    {
                        // WAV header for 24000 Hz, 16-bit mono
                        int sampleRate = 24000;
                        short bitsPerSample = 16;
                        short channels = 1;
                        int byteRate = sampleRate * channels * (bitsPerSample / 8);
                        short blockAlign = (short)(channels * (bitsPerSample / 8));
                        int dataSize = audioBytes.Length;
                        
                        writer.Write(Encoding.UTF8.GetBytes("RIFF"));
                        writer.Write(dataSize + 36);
                        writer.Write(Encoding.UTF8.GetBytes("WAVE"));
                        writer.Write(Encoding.UTF8.GetBytes("fmt "));
                        writer.Write(16);
                        writer.Write((short)1);
                        writer.Write(channels);
                        writer.Write(sampleRate);
                        writer.Write(byteRate);
                        writer.Write(blockAlign);
                        writer.Write(bitsPerSample);
                        writer.Write(Encoding.UTF8.GetBytes("data"));
                        writer.Write(dataSize);
                        writer.Write(audioBytes);
                    }
                    
                    Dispatcher.Invoke(() => LogMessage($"Successfully saved audio to {outputFile}"));
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => LogMessage($"Google TTS Error: {ex.Message}"));
                return false;
            }
        }
        
        // Part 3: SSML Processing and Utility Methods
        private bool ContainsSSMLTags(string text)
        {
            // Check for common SSML tags
            string[] ssmlTags = { "<emphasis", "<break", "<prosody", "<say-as", "<phoneme", "<sub", "<audio", "<p>", "<s>" };
            foreach (var tag in ssmlTags)
            {
                if (text.Contains(tag, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
        
        private string WrapInSSML(string text)
        {
            // If already wrapped in speak tags, return as is
            if (text.TrimStart().StartsWith("<speak", StringComparison.OrdinalIgnoreCase))
                return text;
            
            // Fix pitch values for SAPI (SAPI uses different syntax)
            text = Regex.Replace(text, @"pitch=""\+(\d+)st""", @"pitch=""high""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""-(\d+)st""", @"pitch=""low""", RegexOptions.IgnoreCase);
            
            // Fix date format for SAPI - SAPI prefers different date syntax
            // Convert say-as date to more explicit format for SAPI
            text = PreprocessDatesForSAPI(text);
            
            // Wrap in speak tags for SAPI
            return $"<speak version=\"1.0\" xmlns=\"http://www.w3.org/2001/10/synthesis\" xml:lang=\"en-US\">{text}</speak>";
        }
        
        private string PreprocessDatesForSAPI(string text)
        {
            // Pattern to match say-as date tags with content
            string pattern = @"<say-as\s+interpret-as=[""']date[""'](?:\s+format=[""']([^""']+)[""'])?\s*>([^<]+)</say-as>";
            
            return Regex.Replace(text, pattern, (match) =>
            {
                string format = match.Groups[1].Value;
                string dateContent = match.Groups[2].Value.Trim();
                
                // Try to parse and reformat the date for better SAPI pronunciation
                if (DateTime.TryParse(dateContent, out DateTime date))
                {
                    // Use sub tag for better pronunciation
                    string spokenDate = date.ToString("MMMM d, yyyy");
                    return $"<sub alias=\"{spokenDate}\">{dateContent}</sub>";
                }
                else if (Regex.IsMatch(dateContent, @"^\d{1,2}/\d{1,2}/\d{2,4}$"))
                {
                    // Handle MM/DD/YYYY format
                    var parts = dateContent.Split('/');
                    if (parts.Length == 3)
                    {
                        try
                        {
                            int month = int.Parse(parts[0]);
                            int day = int.Parse(parts[1]);
                            int year = int.Parse(parts[2]);
                            if (year < 100) year += 2000;
                            
                            string[] months = { "", "January", "February", "March", "April", "May", "June",
                                              "July", "August", "September", "October", "November", "December" };
                            
                            if (month >= 1 && month <= 12)
                            {
                                string spokenDate = $"{months[month]} {day}, {year}";
                                return $"<sub alias=\"{spokenDate}\">{dateContent}</sub>";
                            }
                        }
                        catch { }
                    }
                }
                
                // If we can't parse it, return with basic say-as tag
                return $"<say-as interpret-as=\"date\">{dateContent}</say-as>";
            }, RegexOptions.IgnoreCase);
        }
        
        private string ConvertToGoogleSSML(string text)
        {
            // If already wrapped in speak tags, extract the content
            if (text.TrimStart().StartsWith("<speak", StringComparison.OrdinalIgnoreCase))
            {
                // Extract content between speak tags
                var match = Regex.Match(text, @"<speak[^>]*>(.*?)</speak>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (match.Success)
                {
                    text = match.Groups[1].Value;
                }
            }
            
            // Fix common pitch values for Google TTS
            // Google expects semitones (st) or Hz, not words like "high"
            text = Regex.Replace(text, @"pitch=""high""", @"pitch=""+5st""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""x-high""", @"pitch=""+10st""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""low""", @"pitch=""-5st""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""x-low""", @"pitch=""-10st""", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"pitch=""medium""", @"pitch=""+0st""", RegexOptions.IgnoreCase);
            
            // Process dates for Google TTS
            text = PreprocessDatesForGoogle(text);
            
            // Google TTS requires proper SSML format
            return $"<speak>{text}</speak>";
        }
        
        private string PreprocessDatesForGoogle(string text)
        {
            // Pattern to match say-as date tags with content
            string pattern = @"<say-as\s+interpret-as=[""']date[""'](?:\s+format=[""']([^""']+)[""'])?\s*>([^<]+)</say-as>";
            
            return Regex.Replace(text, pattern, (match) =>
            {
                string format = match.Groups[1].Value;
                string dateContent = match.Groups[2].Value.Trim();
                
                // Google TTS handles dates better with specific format attributes
                if (string.IsNullOrEmpty(format))
                {
                    // Try to detect format
                    if (Regex.IsMatch(dateContent, @"^\d{1,2}/\d{1,2}/\d{2,4}$"))
                    {
                        format = "mdy";
                    }
                    else if (Regex.IsMatch(dateContent, @"^\d{4}-\d{2}-\d{2}$"))
                    {
                        format = "ymd";
                    }
                    else if (Regex.IsMatch(dateContent, @"^\d{1,2}-\d{1,2}-\d{2,4}$"))
                    {
                        format = "dmy";
                    }
                }
                
                // Return with explicit format for Google
                if (!string.IsNullOrEmpty(format))
                {
                    return $"<say-as interpret-as=\"date\" format=\"{format}\">{dateContent}</say-as>";
                }
                else
                {
                    // For Google, if we can't determine format, try to parse and use explicit text
                    if (DateTime.TryParse(dateContent, out DateTime date))
                    {
                        string spokenDate = date.ToString("MMMM d, yyyy");
                        return $"<say-as interpret-as=\"date\" format=\"mdy\">{date:MM/dd/yyyy}</say-as>";
                    }
                    return $"<say-as interpret-as=\"date\">{dateContent}</say-as>";
                }
            }, RegexOptions.IgnoreCase);
        }
        
        private void ConvertWavToMp3(string wavFile, string mp3File)
        {
            using (var reader = new AudioFileReader(wavFile))
            {
                using (var writer = new LameMP3FileWriter(mp3File, reader.WaveFormat, LAMEPreset.STANDARD))
                {
                    reader.CopyTo(writer);
                }
            }
        }
        
        // Part 4: Playback Control Methods
        private void UpdatePlaybackControls()
        {
            if (createdAudioFiles.Count > 0)
            {
                cmbCreatedFiles.Items.Clear();
                foreach (var file in createdAudioFiles)
                {
                    cmbCreatedFiles.Items.Add(Path.GetFileName(file));
                }
                
                cmbCreatedFiles.SelectedIndex = 0;
                cmbCreatedFiles.IsEnabled = true;
                btnPlayPause.IsEnabled = true;
                btnStopPlayback.IsEnabled = true;
                btnOpenFolder.IsEnabled = true;
                
                if (createdAudioFiles.Count > 1)
                {
                    btnPlayNext.IsEnabled = true;
                    btnPlayPrevious.IsEnabled = false; // Will enable after first file
                }
            }
        }
        
        private void PlayPause_Click(object sender, RoutedEventArgs e)
        {
            if (createdAudioFiles.Count == 0) return;
            
            if (isPlaying && waveOut != null)
            {
                // Pause
                waveOut.Pause();
                isPlaying = false;
                btnPlayPause.Content = new TextBlock { Text = "▶ Play" };
                playbackTimer.Stop();
            }
            else if (!isPlaying && waveOut != null && waveOut.PlaybackState == PlaybackState.Paused)
            {
                // Resume
                waveOut.Play();
                isPlaying = true;
                btnPlayPause.Content = new TextBlock { Text = "⏸ Pause" };
                playbackTimer.Start();
            }
            else
            {
                // Start new playback
                PlayAudioFile(cmbCreatedFiles.SelectedIndex);
            }
        }
        
        private void StopPlayback_Click(object sender, RoutedEventArgs e)
        {
            StopCurrentPlayback();
        }
        
        private void PlayPrevious_Click(object sender, RoutedEventArgs e)
        {
            if (currentPlayingIndex > 0)
            {
                PlayAudioFile(currentPlayingIndex - 1);
                cmbCreatedFiles.SelectedIndex = currentPlayingIndex;
            }
        }
        
        private void PlayNext_Click(object sender, RoutedEventArgs e)
        {
            if (currentPlayingIndex < createdAudioFiles.Count - 1)
            {
                PlayAudioFile(currentPlayingIndex + 1);
                cmbCreatedFiles.SelectedIndex = currentPlayingIndex;
            }
        }
        
        private void CreatedFile_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (cmbCreatedFiles.SelectedIndex >= 0 && cmbCreatedFiles.SelectedIndex < createdAudioFiles.Count)
            {
                StopCurrentPlayback();
                txtCurrentFile.Text = Path.GetFileName(createdAudioFiles[cmbCreatedFiles.SelectedIndex]);
            }
        }
        
        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(lastOutputFolder) && Directory.Exists(lastOutputFolder))
            {
                System.Diagnostics.Process.Start("explorer.exe", lastOutputFolder);
            }
        }
        
        private void PlayAudioFile(int index)
        {
            if (index < 0 || index >= createdAudioFiles.Count) return;
            
            try
            {
                StopCurrentPlayback();
                
                string filePath = createdAudioFiles[index];
                if (!File.Exists(filePath))
                {
                    LogMessage($"File not found: {filePath}");
                    return;
                }
                
                currentPlayingIndex = index;
                currentAudioFile = new AudioFileReader(filePath);
                waveOut = new WaveOutEvent();
                waveOut.Init(currentAudioFile);
                waveOut.PlaybackStopped += WaveOut_PlaybackStopped;
                waveOut.Play();
                
                isPlaying = true;
                btnPlayPause.Content = new TextBlock { Text = "⏸ Pause" };
                txtCurrentFile.Text = $"Playing: {Path.GetFileName(filePath)}";
                
                // Update navigation buttons
                btnPlayPrevious.IsEnabled = index > 0;
                btnPlayNext.IsEnabled = index < createdAudioFiles.Count - 1;
                
                playbackTimer.Start();
                LogMessage($"Playing: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                LogMessage($"Error playing file: {ex.Message}");
            }
        }
        
        private void StopCurrentPlayback()
        {
            playbackTimer.Stop();
            
            if (waveOut != null)
            {
                waveOut.Stop();
                waveOut.Dispose();
                waveOut = null;
            }
            
            if (currentAudioFile != null)
            {
                currentAudioFile.Dispose();
                currentAudioFile = null;
            }
            
            isPlaying = false;
            btnPlayPause.Content = new TextBlock { Text = "▶ Play" };
            playbackProgress.Value = 0;
            txtPlaybackTime.Text = "00:00 / 00:00";
            txtCurrentFile.Text = "No file playing";
        }
        
        private void WaveOut_PlaybackStopped(object sender, StoppedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                if (isPlaying && currentPlayingIndex < createdAudioFiles.Count - 1)
                {
                    // Auto-play next file
                    PlayAudioFile(currentPlayingIndex + 1);
                    cmbCreatedFiles.SelectedIndex = currentPlayingIndex;
                }
                else
                {
                    StopCurrentPlayback();
                }
            });
        }
        
        private void PlaybackTimer_Tick(object sender, EventArgs e)
        {
            if (currentAudioFile != null && waveOut != null)
            {
                var currentTime = currentAudioFile.CurrentTime;
                var totalTime = currentAudioFile.TotalTime;
                
                playbackProgress.Maximum = totalTime.TotalSeconds;
                playbackProgress.Value = currentTime.TotalSeconds;
                
                txtPlaybackTime.Text = $"{currentTime:mm\\:ss} / {totalTime:mm\\:ss}";
            }
        }
        
        // Part 5: SSML Validation Methods
        private SSMLValidationResult ValidateSSMLSyntax(string text)
        {
            var result = new SSMLValidationResult();
            var tagStack = new Stack<SSMLTag>();
            var selfClosingTags = new HashSet<string> { "break", "phoneme", "audio", "mark", "meta" };
            var validTags = new HashSet<string> 
            { 
                "speak", "emphasis", "break", "prosody", "say-as", "phoneme", "sub", 
                "audio", "p", "s", "voice", "mark", "desc", "lexicon", "metadata", "meta",
                "split", "voice", "service" // Add service to valid tags
            };
            
            // Pattern to match any tag (opening, closing, or self-closing)
            string tagPattern = @"<(/?)([a-zA-Z]+(?:=\d+)?)((?:\s+[a-zA-Z-]+(?:=[""'][^""']*[""'])?)*)\s*(/?)>";
            var matches = Regex.Matches(text, tagPattern);
            
            foreach (Match match in matches)
            {
                string fullTag = match.Value;
                bool isClosing = match.Groups[1].Value == "/";
                string tagName = match.Groups[2].Value.ToLower();
                string attributes = match.Groups[3].Value;
                bool isSelfClosing = match.Groups[4].Value == "/";
                int position = match.Index;
                
                // Handle custom voice tags
                if (tagName.StartsWith("voice="))
                {
                    tagName = "voice";
                    isSelfClosing = true;
                }
                
                // Check if tag is valid
                if (!validTags.Contains(tagName) && !tagName.StartsWith("voice"))
                {
                    result.Errors.Add(new SSMLError
                    {
                        Message = $"Unknown tag: <{tagName}>",
                        Position = position,
                        Length = fullTag.Length,
                        Context = fullTag
                    });
                    continue;
                }
                
                if (isSelfClosing)
                {
                    // Self-closing tag - no need to track
                    if (!selfClosingTags.Contains(tagName) && tagName != "voice" && tagName != "split")
                    {
                        result.Warnings.Add($"Tag <{tagName}> is not typically self-closing");
                    }
                }
                else if (isClosing)
                {
                    // Closing tag - check if it matches the last opened tag
                    if (tagStack.Count == 0)
                    {
                        result.Errors.Add(new SSMLError
                        {
                            Message = $"Closing tag </{tagName}> has no matching opening tag",
                            Position = position,
                            Length = fullTag.Length,
                            Context = fullTag
                        });
                    }
                    else
                    {
                        var expectedTag = tagStack.Pop();
                        if (expectedTag.Name != tagName)
                        {
                            result.Errors.Add(new SSMLError
                            {
                                Message = $"Mismatched closing tag: expected </{expectedTag.Name}>, found </{tagName}>",
                                Position = position,
                                Length = fullTag.Length,
                                Context = $"Opened at position {expectedTag.Position}"
                            });
                            
                            // Push back the expected tag as it wasn't properly closed
                            tagStack.Push(expectedTag);
                        }
                    }
                }
                else
                {
                    // Opening tag
                    if (selfClosingTags.Contains(tagName))
                    {
                        result.Warnings.Add($"Tag <{tagName}> should be self-closing (use <{tagName} />)");
                    }
                    else
                    {
                        tagStack.Push(new SSMLTag { Name = tagName, Position = position });
                        
                        // Validate attributes
                        ValidateTagAttributes(tagName, attributes, position, result);
                    }
                }
            }
            
            // Check for unclosed tags
            while (tagStack.Count > 0)
            {
                var unclosedTag = tagStack.Pop();
                result.Errors.Add(new SSMLError
                {
                    Message = $"Unclosed tag: <{unclosedTag.Name}> at position {unclosedTag.Position}",
                    Position = unclosedTag.Position,
                    Length = unclosedTag.Name.Length + 2,
                    Context = "Tag was never closed"
                });
            }
            
            // Check for proper nesting of special tags
            CheckSpecialTagNesting(text, result);
            
            return result;
        }
        
        private void ValidateTagAttributes(string tagName, string attributes, int position, SSMLValidationResult result)
        {
            // Check required attributes for specific tags
            switch (tagName)
            {
                case "say-as":
                    if (!attributes.Contains("interpret-as"))
                    {
                        result.Errors.Add(new SSMLError
                        {
                            Message = $"<say-as> tag missing required 'interpret-as' attribute",
                            Position = position,
                            Length = tagName.Length + 2,
                            Context = "Required: interpret-as=\"type\""
                        });
                    }
                    break;
                    
                case "break":
                    if (!string.IsNullOrWhiteSpace(attributes) && !attributes.Contains("time") && !attributes.Contains("strength"))
                    {
                        result.Warnings.Add($"<break> tag should have 'time' or 'strength' attribute");
                    }
                    break;
                    
                case "sub":
                    if (!attributes.Contains("alias"))
                    {
                        result.Errors.Add(new SSMLError
                        {
                            Message = $"<sub> tag missing required 'alias' attribute",
                            Position = position,
                            Length = tagName.Length + 2,
                            Context = "Required: alias=\"replacement text\""
                        });
                    }
                    break;
            }
            
            // Check for malformed attributes
            if (!string.IsNullOrWhiteSpace(attributes))
            {
                // Check for unpaired quotes
                int singleQuotes = attributes.Count(c => c == '\'');
                int doubleQuotes = attributes.Count(c => c == '"');
                
                if (singleQuotes % 2 != 0 || doubleQuotes % 2 != 0)
                {
                    result.Errors.Add(new SSMLError
                    {
                        Message = $"Unpaired quotes in attributes for <{tagName}>",
                        Position = position,
                        Length = tagName.Length + attributes.Length + 2,
                        Context = attributes
                    });
                }
            }
        }
        
        private void CheckSpecialTagNesting(string text, SSMLValidationResult result)
        {
            // Check if speak tags are properly used (should be root if present)
            if (text.Contains("<speak", StringComparison.OrdinalIgnoreCase))
            {
                var speakPattern = @"<speak[^>]*>";
                var speakMatches = Regex.Matches(text, speakPattern, RegexOptions.IgnoreCase);
                
                if (speakMatches.Count > 1)
                {
                    result.Warnings.Add("Multiple <speak> tags found. There should only be one root <speak> tag.");
                }
            }
            
            // Check for text outside of tags that might cause issues
            var taglessPattern = @">[^<>]+<";
            var taglessMatches = Regex.Matches(text, taglessPattern);
            
            foreach (Match match in taglessMatches)
            {
                string content = match.Value.Substring(1, match.Value.Length - 2);
                if (content.Contains('&') && !Regex.IsMatch(content, @"&(amp|lt|gt|quot|apos);"))
                {
                    result.Warnings.Add($"Unescaped '&' character found. Use &amp; instead.");
                }
                if (content.Contains('<') || content.Contains('>'))
                {
                    result.Warnings.Add($"Unescaped '<' or '>' found in text. Use &lt; or &gt; instead.");
                }
            }
        }
        
        private void HighlightError(int position, int length)
        {
            try
            {
                var start = GetTextPointerAtOffset(rtbTextContent.Document.ContentStart, position);
                var end = GetTextPointerAtOffset(rtbTextContent.Document.ContentStart, position + length);
                
                if (start != null && end != null)
                {
                    var range = new TextRange(start, end);
                    range.ApplyPropertyValue(TextElement.BackgroundProperty, Brushes.Red);
                    range.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.White);
                    range.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                }
            }
            catch { }
        }
        
        private TextPointer GetTextPointerAtOffset(TextPointer start, int offset)
        {
            var navigator = start;
            int count = 0;
            
            while (navigator != null && count < offset)
            {
                if (navigator.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                {
                    int remainingLength = offset - count;
                    string textRun = navigator.GetTextInRun(LogicalDirection.Forward);
                    int textLength = textRun.Length;
                    
                    if (textLength >= remainingLength)
                    {
                        return navigator.GetPositionAtOffset(remainingLength);
                    }
                    
                    count += textLength;
                }
                
                navigator = navigator.GetNextContextPosition(LogicalDirection.Forward);
            }
            
            return navigator;
        }
        
        private void LogMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                txtLog.ScrollToEnd();
            });
        }
        
        // Helper classes
        private class TextSegment
        {
            public string Text { get; set; }
            public int VoiceIndex { get; set; }
            public int ServiceIndex { get; set; } = -1;
            public int SplitIndex { get; set; } = 0;  // Which <split> section
            public int SubIndex { get; set; } = 0;    // Sub-segment within split
        }
        
        private class SSMLValidationResult
        {
            public List<SSMLError> Errors { get; set; } = new List<SSMLError>();
            public List<string> Warnings { get; set; } = new List<string>();
        }
        
        private class SSMLError
        {
            public string Message { get; set; }
            public int Position { get; set; }
            public int Length { get; set; }
            public string Context { get; set; }
        }
        
        private class SSMLTag
        {
            public string Name { get; set; }
            public int Position { get; set; }
        }
        
        // Unused legacy methods (kept for compatibility)
        private async Task ConvertWithSAPI(TextSegment segment, string outputFile)
        {
            await ConvertWithSAPIBackground(segment, outputFile, sliderRate.Value, sliderVolume.Value, cmbOutputFormat.SelectedIndex);
        }
        
        private async Task ConvertWithGoogleTTS(TextSegment segment, string outputFile)
        {
            await ConvertWithGoogleTTSBackground(segment, outputFile, sliderRate.Value, sliderVolume.Value, cmbOutputFormat.SelectedIndex);
        }
		// Add these methods to the MainWindow class:

		private void Exit_Click(object sender, RoutedEventArgs e)
		{
			// Clean up resources before exiting
			StopCurrentPlayback();
			
			if (speechSynthesizer != null)
			{
				speechSynthesizer.Dispose();
			}
			
			if (httpClient != null)
			{
				httpClient.Dispose();
			}
                if (pollyClient != null)
    {
        pollyClient.Dispose();
    }
    
    Application.Current.Shutdown();
			
			Application.Current.Shutdown();
		}

		private void About_Click(object sender, RoutedEventArgs e)
		{
			var aboutWindow = new Window
			{
				Title = "About Text to Speech Converter",
				Width = 400,
				Height = 250,
				WindowStartupLocation = WindowStartupLocation.CenterOwner,
				Owner = this,
				ResizeMode = ResizeMode.NoResize
			};
			
			var grid = new Grid();
			grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			
			var stackPanel = new StackPanel
			{
				Margin = new Thickness(20),
				VerticalAlignment = VerticalAlignment.Center
			};
			
			stackPanel.Children.Add(new TextBlock
			{
				Text = "Text to Speech Converter",
				FontSize = 20,
				FontWeight = FontWeights.Bold,
				HorizontalAlignment = HorizontalAlignment.Center,
				Margin = new Thickness(0, 0, 0, 10)
			});
			
			stackPanel.Children.Add(new TextBlock
			{
				Text = "Version 1.0",
				HorizontalAlignment = HorizontalAlignment.Center,
				Margin = new Thickness(0, 0, 0, 20)
			});
			
			stackPanel.Children.Add(new TextBlock
			{
				Text = "Designed by M. Sodhi",
				HorizontalAlignment = HorizontalAlignment.Center,
				FontSize = 14,
				Margin = new Thickness(0, 0, 0, 5)
			});
			
			var emailLink = new TextBlock
			{
				HorizontalAlignment = HorizontalAlignment.Center,
				FontSize = 14,
				Margin = new Thickness(0, 0, 0, 20)
			};
			
			var hyperlink = new System.Windows.Documents.Hyperlink(new System.Windows.Documents.Run("sodhi@minmaxsolutions.com"))
			{
				NavigateUri = new Uri("mailto:sodhi@minmaxsolutions.com")
			};
			hyperlink.RequestNavigate += (s, args) =>
			{
				System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
				{
					FileName = args.Uri.ToString(),
					UseShellExecute = true
				});
				args.Handled = true;
			};
			
			emailLink.Inlines.Add(hyperlink);
			stackPanel.Children.Add(emailLink);
			
			stackPanel.Children.Add(new TextBlock
			{
				Text = "© 2024 MinMax Solutions",
				HorizontalAlignment = HorizontalAlignment.Center,
				FontSize = 12,
				Foreground = Brushes.Gray
			});
			
			var closeButton = new Button
			{
				Content = "Close",
				Width = 100,
				Height = 30,
				Margin = new Thickness(10),
				HorizontalAlignment = HorizontalAlignment.Center
			};
			
			closeButton.Click += (s, args) => aboutWindow.Close();
			
			Grid.SetRow(stackPanel, 0);
			Grid.SetRow(closeButton, 1);
			
			grid.Children.Add(stackPanel);
			grid.Children.Add(closeButton);
			
			aboutWindow.Content = grid;
			aboutWindow.ShowDialog();
		}

		// Updated APIKey_Click method with proper label
		private void APIKey_Click(object sender, RoutedEventArgs e)
		{
			var dialog = new Window
			{
				Title = "Enter Google Cloud API Key",
				Width = 450,
				Height = 180,
				WindowStartupLocation = WindowStartupLocation.CenterOwner,
				Owner = this,
				ResizeMode = ResizeMode.NoResize
			};
			
			var grid = new Grid();
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			
			var labelText = new TextBlock
			{
				Text = "Enter your Google Cloud Text-to-Speech API Key:",
				Margin = new Thickness(10, 10, 10, 5),
				FontWeight = FontWeights.SemiBold
			};
			
			var instructionText = new TextBlock
			{
				Text = "You can get an API key from the Google Cloud Console",
				Margin = new Thickness(10, 0, 10, 10),
				FontSize = 11,
				Foreground = Brushes.Gray
			};
			
			var textBox = new TextBox 
			{ 
				Margin = new Thickness(10),
				VerticalAlignment = VerticalAlignment.Center,
				Text = googleApiKey,
				FontFamily = new FontFamily("Consolas")
			};
			
			var buttonPanel = new StackPanel
			{
				Orientation = Orientation.Horizontal,
				HorizontalAlignment = HorizontalAlignment.Right,
				Margin = new Thickness(10)
			};
			
			var okButton = new Button { Content = "OK", Width = 75, Margin = new Thickness(5) };
			var cancelButton = new Button { Content = "Cancel", Width = 75, Margin = new Thickness(5) };
			
            okButton.Click += (s, args) =>
            {
                googleApiKey = textBox.Text.Trim();
                if (string.IsNullOrEmpty(googleApiKey))
                {
                    LogMessage("Google API Key cleared");
                }
                else
                {
                    LogMessage("Google API Key set successfully");
                }
                
                // Save credentials
                SaveCredentials();
                
                dialog.DialogResult = true;
            };			
			cancelButton.Click += (s, args) => dialog.DialogResult = false;
			
			buttonPanel.Children.Add(okButton);
			buttonPanel.Children.Add(cancelButton);
			
			Grid.SetRow(labelText, 0);
			Grid.SetRow(instructionText, 1);
			Grid.SetRow(textBox, 2);
			Grid.SetRow(buttonPanel, 3);
			
			grid.Children.Add(labelText);
			grid.Children.Add(instructionText);
			grid.Children.Add(textBox);
			grid.Children.Add(buttonPanel);
			
			dialog.Content = grid;
			
			// Focus on textbox when dialog opens
			dialog.Loaded += (s, args) => 
			{
				textBox.Focus();
				textBox.SelectAll();
			};
			
			dialog.ShowDialog();
		}
        private void ElevenLabsKey_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Window
            {
                Title = "Enter ElevenLabs API Key",
                Width = 450,
                Height = 200,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize
            };
            
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            
            var labelText = new TextBlock
            {
                Text = "Enter your ElevenLabs API Key:",
                Margin = new Thickness(10, 10, 10, 5),
                FontWeight = FontWeights.SemiBold
            };
            
            var instructionText = new TextBlock
            {
                Text = "Get your API key from https://elevenlabs.io/speech-synthesis",
                Margin = new Thickness(10, 0, 10, 10),
                FontSize = 11,
                Foreground = Brushes.Gray
            };
            
            var textBox = new TextBox 
            { 
                Margin = new Thickness(10),
                VerticalAlignment = VerticalAlignment.Center,
                Text = elevenLabsApiKey,
                FontFamily = new FontFamily("Consolas")
            };
            
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(10)
            };
            
            var okButton = new Button { Content = "OK", Width = 75, Margin = new Thickness(5) };
            var cancelButton = new Button { Content = "Cancel", Width = 75, Margin = new Thickness(5) };
            
            okButton.Click += (s, args) =>
            {
                elevenLabsApiKey = textBox.Text.Trim();
                if (string.IsNullOrEmpty(elevenLabsApiKey))
                {
                    LogMessage("ElevenLabs API Key cleared");
                }
                else
                {
                    LogMessage("ElevenLabs API Key set successfully");
                }
                SaveCredentials();
                dialog.DialogResult = true;
            };
            
            cancelButton.Click += (s, args) => dialog.DialogResult = false;
            
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            
            Grid.SetRow(labelText, 0);
            Grid.SetRow(instructionText, 1);
            Grid.SetRow(textBox, 2);
            Grid.SetRow(buttonPanel, 3);
            
            grid.Children.Add(labelText);
            grid.Children.Add(instructionText);
            grid.Children.Add(textBox);
            grid.Children.Add(buttonPanel);
            
            dialog.Content = grid;
            
            dialog.Loaded += (s, args) => 
            {
                textBox.Focus();
                textBox.SelectAll();
            };
            
            dialog.ShowDialog();
        }

        private async Task<bool> CallElevenLabs(string text, string outputFile)
        {
            try
            {
                if (string.IsNullOrEmpty(elevenLabsApiKey))
                {
                    LogMessage("ElevenLabs API key not set");
                    return false;
                }
                
                string url = $"https://api.elevenlabs.io/v1/text-to-speech/{currentElevenLabsVoice}";
                
                var requestBody = new
                {
                    text = text,
                    model_id = "eleven_monolingual_v1",
                    voice_settings = new
                    {
                        stability = 0.5,
                        similarity_boost = 0.75,
                        style = 0.0,
                        use_speaker_boost = true
                    }
                };
                
                var json = JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Headers.Add("xi-api-key", elevenLabsApiKey);
                    request.Content = content;
                    
                    LogMessage($"Calling ElevenLabs with voice ID: {currentElevenLabsVoice}");
                    
                    var response = await httpClient.SendAsync(request);
                    
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        LogMessage($"ElevenLabs Error {response.StatusCode}: {errorContent}");
                        return false;
                    }
                    
                    var audioBytes = await response.Content.ReadAsByteArrayAsync();
                    
                    // ElevenLabs returns MP3, convert to WAV
                    string tempMp3 = Path.GetTempFileName() + ".mp3";
                    File.WriteAllBytes(tempMp3, audioBytes);
                    
                    try
                    {
                        // Convert MP3 to WAV
                        using (var reader = new Mp3FileReader(tempMp3))
                        using (var writer = new WaveFileWriter(outputFile, reader.WaveFormat))
                        {
                            reader.CopyTo(writer);
                        }
                        
                        File.Delete(tempMp3);
                        LogMessage($"Successfully saved audio to {outputFile}");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Error converting MP3 to WAV: {ex.Message}");
                        File.Delete(tempMp3);
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ElevenLabs Error: {ex.Message}");
                return false;
            }
        }

        private async Task<bool> CallElevenLabsWithSettings(string text, string outputFile, 
            string voiceToUse, double rateValue, double volumeValue)
        {
            try
            {
                if (string.IsNullOrEmpty(elevenLabsApiKey))
                {
                    throw new Exception("ElevenLabs API key not set");
                }
                
                string url = $"https://api.elevenlabs.io/v1/text-to-speech/{voiceToUse}";
                
                // ElevenLabs doesn't have direct rate/volume controls like other services
                // Stability and similarity_boost are the main controls
                var requestBody = new
                {
                    text = text,
                    model_id = "eleven_monolingual_v1",
                    voice_settings = new
                    {
                        stability = 0.5 + (rateValue / 20.0), // Adjust based on rate
                        similarity_boost = 0.75,
                        style = 0.0,
                        use_speaker_boost = true
                    }
                };
                
                var json = JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Headers.Add("xi-api-key", elevenLabsApiKey);
                    request.Content = content;
                    
                    Dispatcher.Invoke(() => LogMessage($"Calling ElevenLabs with voice ID: {voiceToUse}"));
                    
                    var response = await httpClient.SendAsync(request);
                    
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        Dispatcher.Invoke(() => LogMessage($"ElevenLabs Error {response.StatusCode}: {errorContent}"));
                        return false;
                    }
                    
                    var audioBytes = await response.Content.ReadAsByteArrayAsync();
                    
                    // ElevenLabs returns MP3 by default, convert to WAV if needed
                    string tempMp3 = Path.GetTempFileName() + ".mp3";
                    File.WriteAllBytes(tempMp3, audioBytes);
                    
                    // Convert MP3 to WAV
                    using (var reader = new Mp3FileReader(tempMp3))
                    using (var writer = new WaveFileWriter(outputFile, reader.WaveFormat))
                    {
                        reader.CopyTo(writer);
                    }
                    
                    File.Delete(tempMp3);
                    
                    Dispatcher.Invoke(() => LogMessage($"Successfully saved audio to {outputFile}"));
                    return true;
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => LogMessage($"ElevenLabs Error: {ex.Message}"));
                return false;
            }
        }

        private async Task ConvertWithElevenLabsBackground(TextSegment segment, string outputFile,
            double rateValue, double volumeValue, int outputFormatIndex)
        {
            if (string.IsNullOrEmpty(elevenLabsApiKey))
            {
                throw new Exception("ElevenLabs API key not set");
            }
            
            var voicesList = elevenLabsVoices.Values.ToList();
            
            // Default to first voice
            string voiceToUse = voicesList.Count > 0 ? voicesList[0] : "21m00Tcm4TlvDq8ikWAM";
            
            if (segment.VoiceIndex >= 0 && segment.VoiceIndex < voicesList.Count)
            {
                voiceToUse = voicesList[segment.VoiceIndex];
                Dispatcher.Invoke(() => 
                    LogMessage($"ElevenLabs: Using voice ID {voiceToUse} (index {segment.VoiceIndex})"));
            }
            else
            {
                Dispatcher.Invoke(() => 
                    LogMessage($"ElevenLabs: Using default voice ID {voiceToUse}"));
            }
            
            string extension = outputFormatIndex == 0 ? ".wav" : ".mp3";
            string wavFile = outputFile + ".wav";
            
            bool success = await CallElevenLabsWithSettings(segment.Text, wavFile, voiceToUse, rateValue, volumeValue);
            
            if (!success)
            {
                throw new Exception("ElevenLabs API call failed");
            }
            
            if (extension == ".mp3")
            {
                ConvertWavToMp3(wavFile, outputFile + ".mp3");
                File.Delete(wavFile);
            }
        }
		private void FindReplace_Click(object sender, RoutedEventArgs e)
		{
			var findReplaceWindow = new Window
			{
				Title = "Find and Replace",
				Width = 500,
				Height = 350,
				WindowStartupLocation = WindowStartupLocation.CenterOwner,
				Owner = this,
				ResizeMode = ResizeMode.NoResize
			};
			
			var grid = new Grid();
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
			grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
			
			// Find section
			var findLabel = new Label { Content = "Find what:", Margin = new Thickness(10, 10, 10, 0) };
			var findTextBox = new TextBox 
			{ 
				Margin = new Thickness(10, 5, 10, 10),
				Height = 25,
				FontFamily = new FontFamily("Consolas")
			};
			
			// Replace section
			var replaceLabel = new Label { Content = "Replace with:", Margin = new Thickness(10, 0, 10, 0) };
			var replaceTextBox = new TextBox 
			{ 
				Margin = new Thickness(10, 5, 10, 10),
				Height = 25,
				FontFamily = new FontFamily("Consolas")
			};
			
			// Quick replacements dropdown
			var quickLabel = new Label { Content = "Quick replacements:", Margin = new Thickness(10, 0, 10, 0) };
			var quickReplacementsCombo = new ComboBox
			{
				Margin = new Thickness(10, 5, 10, 10),
				Height = 25
			};
			
			// Add quick replacement options
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "Select a quick replacement..." });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "Slide → <split>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "SLIDE → <split>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[pause] → <break time=\"1s\"/>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[PAUSE] → <break time=\"1s\"/>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[emphasis] → <emphasis level=\"strong\">" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[/emphasis] → </emphasis>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[slow] → <prosody rate=\"slow\">" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[/slow] → </prosody>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[fast] → <prosody rate=\"fast\">" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[/fast] → </prosody>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[voice1] → <voice=1>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[voice2] → <voice=2>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[voice3] → <voice=3>" });
			quickReplacementsCombo.Items.Add(new ComboBoxItem { Content = "[voice4] → <voice=4>" });
			
			quickReplacementsCombo.SelectedIndex = 0;
			quickReplacementsCombo.SelectionChanged += (s, args) =>
			{
				if (quickReplacementsCombo.SelectedIndex > 0)
				{
					var selectedItem = (quickReplacementsCombo.SelectedItem as ComboBoxItem)?.Content.ToString();
					if (!string.IsNullOrEmpty(selectedItem) && selectedItem.Contains(" → "))
					{
						var parts = selectedItem.Split(new[] { " → " }, StringSplitOptions.None);
						if (parts.Length == 2)
						{
							findTextBox.Text = parts[0];
							replaceTextBox.Text = parts[1];
						}
					}
				}
			};
			
			// Options
			var optionsPanel = new StackPanel
			{
				Orientation = Orientation.Horizontal,
				Margin = new Thickness(10, 0, 10, 10)
			};
			
			var matchCaseCheckBox = new CheckBox 
			{ 
				Content = "Match case",
				Margin = new Thickness(0, 0, 20, 0),
				VerticalAlignment = VerticalAlignment.Center
			};
			
			var wholeWordCheckBox = new CheckBox 
			{ 
				Content = "Match whole word only",
				VerticalAlignment = VerticalAlignment.Center
			};
			
			optionsPanel.Children.Add(matchCaseCheckBox);
			optionsPanel.Children.Add(wholeWordCheckBox);
			
			// Statistics label
			var statsLabel = new Label
			{
				Content = "",
				Margin = new Thickness(10, 0, 10, 10),
				FontStyle = FontStyles.Italic
			};
			
			// Button panel
			var buttonPanel = new StackPanel
			{
				Orientation = Orientation.Horizontal,
				HorizontalAlignment = HorizontalAlignment.Center,
				Margin = new Thickness(10)
			};
			
			var findNextButton = new Button 
			{ 
				Content = "Find Next",
				Width = 100,
				Height = 30,
				Margin = new Thickness(5)
			};
			
			var replaceButton = new Button 
			{ 
				Content = "Replace",
				Width = 100,
				Height = 30,
				Margin = new Thickness(5)
			};
			
			var replaceAllButton = new Button 
			{ 
				Content = "Replace All",
				Width = 100,
				Height = 30,
				Margin = new Thickness(5)
			};
			
			var closeButton = new Button 
			{ 
				Content = "Close",
				Width = 100,
				Height = 30,
				Margin = new Thickness(5)
			};
			
			// Current find position tracking
			TextPointer currentFindPosition = rtbTextContent.Document.ContentStart;
			
			// Find Next functionality
			findNextButton.Click += (s, args) =>
			{
				string searchText = findTextBox.Text;
				if (string.IsNullOrEmpty(searchText)) return;
				
				var textRange = new TextRange(rtbTextContent.Document.ContentStart, rtbTextContent.Document.ContentEnd);
				string documentText = textRange.Text;
				
				// Start searching from current position
				var searchStart = currentFindPosition.GetOffsetToPosition(rtbTextContent.Document.ContentStart);
				
				StringComparison comparison = matchCaseCheckBox.IsChecked == true ? 
					StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
				
				int index = documentText.IndexOf(searchText, Math.Abs(searchStart), comparison);
				
				if (index == -1 && searchStart != 0)
				{
					// Wrap around to beginning
					index = documentText.IndexOf(searchText, 0, comparison);
				}
				
				if (index >= 0)
				{
					var start = GetTextPointerAtOffset(rtbTextContent.Document.ContentStart, index);
					var end = GetTextPointerAtOffset(rtbTextContent.Document.ContentStart, index + searchText.Length);
					
					if (start != null && end != null)
					{
						rtbTextContent.Selection.Select(start, end);
						rtbTextContent.Focus();
						currentFindPosition = end;
					}
				}
				else
				{
					MessageBox.Show("No more occurrences found.", "Find", MessageBoxButton.OK, MessageBoxImage.Information);
					currentFindPosition = rtbTextContent.Document.ContentStart;
				}
			};
			
			// Replace functionality
			replaceButton.Click += (s, args) =>
			{
				if (!rtbTextContent.Selection.IsEmpty && rtbTextContent.Selection.Text == findTextBox.Text)
				{
					rtbTextContent.Selection.Text = replaceTextBox.Text;
				}
				
				// Find next occurrence
				findNextButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
			};
			
			// Replace All functionality
			replaceAllButton.Click += (s, args) =>
			{
				string findText = findTextBox.Text;
				string replaceText = replaceTextBox.Text;
				
				if (string.IsNullOrEmpty(findText))
				{
					MessageBox.Show("Please enter text to find.", "Replace All", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				
				var textRange = new TextRange(rtbTextContent.Document.ContentStart, rtbTextContent.Document.ContentEnd);
				string content = textRange.Text;
				
				StringComparison comparison = matchCaseCheckBox.IsChecked == true ? 
					StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
				
				int count = 0;
				int index = 0;
				
				// Count occurrences
				while ((index = content.IndexOf(findText, index, comparison)) != -1)
				{
					count++;
					index += findText.Length;
				}
				
				if (count == 0)
				{
					MessageBox.Show("No occurrences found.", "Replace All", MessageBoxButton.OK, MessageBoxImage.Information);
					return;
				}
				
				// Confirm replacement
				var result = MessageBox.Show(
					$"Replace {count} occurrence(s) of '{findText}' with '{replaceText}'?",
					"Confirm Replace All",
					MessageBoxButton.YesNo,
					MessageBoxImage.Question);
				
				if (result == MessageBoxResult.Yes)
				{
					// Perform replacement
					if (matchCaseCheckBox.IsChecked == true)
					{
						content = content.Replace(findText, replaceText);
					}
					else
					{
						content = Regex.Replace(content, 
							Regex.Escape(findText), 
							replaceText.Replace("$", "$$"), 
							RegexOptions.IgnoreCase);
					}
					
					// Update the document
					rtbTextContent.Document.Blocks.Clear();
					rtbTextContent.Document.Blocks.Add(new Paragraph(new Run(content)));
					
					statsLabel.Content = $"Replaced {count} occurrence(s)";
					LogMessage($"Replaced {count} occurrence(s) of '{findText}' with '{replaceText}'");
				}
			};
			
			closeButton.Click += (s, args) => findReplaceWindow.Close();
			
			// Update statistics when find text changes
			findTextBox.TextChanged += (s, args) =>
			{
				string searchText = findTextBox.Text;
				if (string.IsNullOrEmpty(searchText))
				{
					statsLabel.Content = "";
					return;
				}
				
				var textRange = new TextRange(rtbTextContent.Document.ContentStart, rtbTextContent.Document.ContentEnd);
				string content = textRange.Text;
				
				StringComparison comparison = matchCaseCheckBox.IsChecked == true ? 
					StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
				
				int count = 0;
				int index = 0;
				
				while ((index = content.IndexOf(searchText, index, comparison)) != -1)
				{
					count++;
					index += searchText.Length;
				}
				
				statsLabel.Content = count > 0 ? $"Found {count} occurrence(s)" : "No occurrences found";
			};
			
			// Layout
			Grid.SetRow(findLabel, 0);
			Grid.SetRow(findTextBox, 1);
			Grid.SetRow(replaceLabel, 2);
			Grid.SetRow(replaceTextBox, 3);
			Grid.SetRow(quickLabel, 4);
			Grid.SetRow(quickReplacementsCombo, 5);
			Grid.SetRow(optionsPanel, 6);
			Grid.SetRow(statsLabel, 7);
			Grid.SetRow(buttonPanel, 8);
			
			grid.Children.Add(findLabel);
			grid.Children.Add(findTextBox);
			grid.Children.Add(replaceLabel);
			grid.Children.Add(replaceTextBox);
			grid.Children.Add(quickLabel);
			grid.Children.Add(quickReplacementsCombo);
			grid.Children.Add(optionsPanel);
			grid.Children.Add(statsLabel);
			grid.Children.Add(buttonPanel);
			
			buttonPanel.Children.Add(findNextButton);
			buttonPanel.Children.Add(replaceButton);
			buttonPanel.Children.Add(replaceAllButton);
			buttonPanel.Children.Add(closeButton);
			
			findReplaceWindow.Content = grid;
			
			// Focus on find textbox when window opens
			findReplaceWindow.Loaded += (s, args) => 
			{
				findTextBox.Focus();
				
				// If there's selected text, use it as the find text
				if (!rtbTextContent.Selection.IsEmpty)
				{
					findTextBox.Text = rtbTextContent.Selection.Text;
					findTextBox.SelectAll();
				}
			};
			
			// Handle keyboard shortcuts in the window
			findReplaceWindow.PreviewKeyDown += (s, args) =>
			{
				if (args.Key == Key.Escape)
				{
					findReplaceWindow.Close();
				}
				else if (args.Key == Key.F3 || (args.Key == Key.F && Keyboard.Modifiers == ModifierKeys.Control))
				{
					findNextButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
					args.Handled = true;
				}
				else if (args.Key == Key.H && Keyboard.Modifiers == ModifierKeys.Control)
				{
					replaceButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
					args.Handled = true;
				}
			};
			
			findReplaceWindow.ShowDialog();
		}
		private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
		{
			// Ctrl+H for Find/Replace
			if (e.Key == Key.H && Keyboard.Modifiers == ModifierKeys.Control)
			{
				FindReplace_Click(sender, e);
				e.Handled = true;
			}
			// Ctrl+O for Open
			else if (e.Key == Key.O && Keyboard.Modifiers == ModifierKeys.Control)
			{
				OpenFile_Click(sender, e);
				e.Handled = true;
			}
			// Ctrl+S for Save
			else if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
			{
				SaveText_Click(sender, e);
				e.Handled = true;
			}
			// F5 for Convert/Generate Audio
			else if (e.Key == Key.F5)
			{
				if (btnConvert.IsEnabled)
				{
					Convert_Click(sender, e);
					e.Handled = true;
				}
			}
			// Space for Play/Pause when focus is not in text editor
			else if (e.Key == Key.Space && !rtbTextContent.IsFocused && btnPlayPause.IsEnabled)
			{
				PlayPause_Click(sender, e);
				e.Handled = true;
			}
		}
    }
}
