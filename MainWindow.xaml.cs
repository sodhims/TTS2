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
using Microsoft.Win32;
using NAudio.Wave;
using NAudio.Lame;
using System.Net.Http;
using System.Text.Json;

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
            else // Google TTS
            {
                // Add Google voices
                foreach (var voice in googleVoices.Keys)
                {
                    cmbVoices.Items.Add(voice);
                }
            }
            
            if (cmbVoices.Items.Count > 0)
                cmbVoices.SelectedIndex = 0;
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
            if (btnAPIKey != null)
            {
                btnAPIKey.Visibility = cmbTTSEngine.SelectedIndex == 1 ? 
                    Visibility.Visible : Visibility.Collapsed;
            }
            LoadVoices();
        }
        
        private void APIKey_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Window
            {
                Title = "Enter Google Cloud API Key",
                Width = 400,
                Height = 150,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this
            };
            
            var grid = new Grid();
            var textBox = new TextBox 
            { 
                Margin = new Thickness(10),
                VerticalAlignment = VerticalAlignment.Center,
                Text = googleApiKey
            };
            
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(10)
            };
            
            var okButton = new Button { Content = "OK", Width = 75, Margin = new Thickness(5) };
            var cancelButton = new Button { Content = "Cancel", Width = 75, Margin = new Thickness(5) };
            
            okButton.Click += (s, args) =>
            {
                googleApiKey = textBox.Text;
                LogMessage("Google API Key set successfully");
                dialog.DialogResult = true;
            };
            
            cancelButton.Click += (s, args) => dialog.DialogResult = false;
            
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            
            grid.Children.Add(textBox);
            grid.Children.Add(buttonPanel);
            dialog.Content = grid;
            
            dialog.ShowDialog();
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
                
                var requestBody = new
                {
                    input = new { text = text },
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
        
        private void InsertSplit_Click(object sender, RoutedEventArgs e)
        {
            var caretPos = rtbTextContent.CaretPosition;
            caretPos.InsertTextInRun("<split></split>");
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
                caretPos.InsertTextInRun(ssmlTag);
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
                else // Google TTS
                {
                    if (string.IsNullOrEmpty(googleApiKey))
                    {
                        MessageBox.Show("Please set Google API key first");
                        return;
                    }
                    
                    string tempFile = Path.GetTempFileName() + ".wav";
                    bool success = await CallGoogleTTS(testText, tempFile);
                    
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
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error testing voice: {ex.Message}");
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
                    
                    for (int i = 0; i < segments.Count; i++)
                    {
                        // Update progress on UI thread
                        Dispatcher.Invoke(() => 
                        {
                            progressBar.Value = (i * 100) / segments.Count;
                        });
                        
                        string outputFile = Path.Combine(outputPath, $"output_{i + 1:D3}");
                        
                        Dispatcher.Invoke(() => 
                            LogMessage($"Segment {i+1}: Voice={segments[i].VoiceIndex}, Length={segments[i].Text.Length} chars"));
                        
                        if (engineIndex == 0)
                        {
                            await ConvertWithSAPIBackground(segments[i], outputFile, rateValue, volumeValue, outputFormatIndex);
                        }
                        else
                        {
                            await ConvertWithGoogleTTSBackground(segments[i], outputFile, rateValue, volumeValue, outputFormatIndex);
                        }
                        
                        string extension = outputFormatIndex == 0 ? ".wav" : ".mp3";
                        string fullPath = outputFile + extension;
                        createdAudioFiles.Add(fullPath);
                        
                        Dispatcher.Invoke(() => LogMessage($"Created: {fullPath}"));
                    }
                    
                    Dispatcher.Invoke(() =>
                    {
                        progressBar.Visibility = Visibility.Collapsed;
                        txtStatus.Text = "Conversion complete";
                        LogMessage($"All {segments.Count} files converted successfully!");
                        
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
            
            // Wrap in speak tags for SAPI
            return $"<speak version=\"1.0\" xmlns=\"http://www.w3.org/2001/10/synthesis\" xml:lang=\"en-US\">{text}</speak>";
        }
        
        private async Task ConvertWithGoogleTTSBackground(TextSegment segment, string outputFile,
            double rateValue, double volumeValue, int outputFormatIndex)
        {
            if (string.IsNullOrEmpty(googleApiKey))
            {
                throw new Exception("Google API key not set");
            }
            
            // Get available voices
            var voicesList = googleVoices.Values.ToList();
            
            // Select the appropriate voice based on segment's voice index
            string voiceToUse = currentGoogleVoice;
            if (segment.VoiceIndex >= 0 && segment.VoiceIndex < voicesList.Count)
            {
                voiceToUse = voicesList[segment.VoiceIndex];
                Dispatcher.Invoke(() => 
                    LogMessage($"Google TTS: Using voice {voiceToUse} (index {segment.VoiceIndex})"));
            }
            
            string extension = outputFormatIndex == 0 ? ".wav" : ".mp3";
            string wavFile = outputFile + ".wav";
            
            bool success = await CallGoogleTTSWithSettings(segment.Text, wavFile, voiceToUse, rateValue, volumeValue);
            
            if (success && extension == ".mp3")
            {
                ConvertWavToMp3(wavFile, outputFile + ".mp3");
                File.Delete(wavFile);
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
                
                var requestBody = new
                {
                    input = new { text = text },
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
        
        private List<TextSegment> ProcessTextSegments(string text)
        {
            var segments = new List<TextSegment>();
            
            // First, split by <split> tags
            var splitPattern = @"<split>";
            var parts = Regex.Split(text, splitPattern, RegexOptions.IgnoreCase);
            
            int currentVoiceIndex = 0;
            
            foreach (var part in parts)
            {
                if (string.IsNullOrWhiteSpace(part)) continue;
                
                // Now process voice changes within each part
                var voicePattern = @"<voice=(\d+)>";
                var voiceMatches = Regex.Matches(part, voicePattern);
                
                if (voiceMatches.Count == 0)
                {
                    // No voice changes in this segment
                    var cleanText = part.Trim();
                    if (!string.IsNullOrEmpty(cleanText))
                    {
                        segments.Add(new TextSegment 
                        { 
                            Text = cleanText, 
                            VoiceIndex = currentVoiceIndex 
                        });
                    }
                }
                else
                {
                    // Process text with voice changes
                    int lastIndex = 0;
                    
                    foreach (Match match in voiceMatches)
                    {
                        // Add text before the voice tag
                        if (match.Index > lastIndex)
                        {
                            var beforeText = part.Substring(lastIndex, match.Index - lastIndex).Trim();
                            if (!string.IsNullOrEmpty(beforeText))
                            {
                                segments.Add(new TextSegment 
                                { 
                                    Text = beforeText, 
                                    VoiceIndex = currentVoiceIndex 
                                });
                            }
                        }
                        
                        // Update voice index
                        currentVoiceIndex = int.Parse(match.Groups[1].Value) - 1;
                        lastIndex = match.Index + match.Length;
                    }
                    
                    // Add remaining text after last voice tag
                    if (lastIndex < part.Length)
                    {
                        var remainingText = part.Substring(lastIndex).Trim();
                        if (!string.IsNullOrEmpty(remainingText))
                        {
                            segments.Add(new TextSegment 
                            { 
                                Text = remainingText, 
                                VoiceIndex = currentVoiceIndex 
                            });
                        }
                    }
                }
            }
            
            if (segments.Count == 0)
            {
                segments.Add(new TextSegment { Text = text, VoiceIndex = 0 });
            }
            
            return segments;
        }
        
        private async Task ConvertWithSAPI(TextSegment segment, string outputFile)
        {
            // This method is no longer used - replaced by ConvertWithSAPIBackground
            await ConvertWithSAPIBackground(segment, outputFile, sliderRate.Value, sliderVolume.Value, cmbOutputFormat.SelectedIndex);
        }
        
        private async Task ConvertWithGoogleTTS(TextSegment segment, string outputFile)
        {
            // This method is no longer used - replaced by ConvertWithGoogleTTSBackground  
            await ConvertWithGoogleTTSBackground(segment, outputFile, sliderRate.Value, sliderVolume.Value, cmbOutputFormat.SelectedIndex);
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
        
        private void LogMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                txtLog.ScrollToEnd();
            });
        }
        
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
        
        private class TextSegment
        {
            public string Text { get; set; }
            public int VoiceIndex { get; set; }
        }
    }
}
