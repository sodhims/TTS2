# Text to Speech Converter

A WPF desktop application for converting text to speech using multiple TTS engines.

## Features

- Support for 3 TTS engines:
  - Windows SAPI (built-in)
  - Google Cloud Text-to-Speech
  - AWS Polly
- SSML support for advanced speech control
- Multi-voice support with voice tags
- Split documents into multiple audio files
- Export to WAV or MP3 format
- Find and replace with quick SSML replacements
- SSML validation and highlighting

## Setup

### Prerequisites

- .NET 8.0 or later
- Windows OS

### Configuration

1. Copy `App.config.example` to `App.config`
2. Enter your API credentials in the application:
   - **Google Cloud TTS**: Click "Set API Key" button
   - **AWS Polly**: Click "AWS Credentials" button

### AWS Polly Setup

1. Create an IAM user in AWS Console
2. Attach the `AmazonPollyFullAccess` policy
3. Create access keys and enter them in the app

### Google Cloud TTS Setup

1. Enable Cloud Text-to-Speech API in Google Cloud Console
2. Create an API key
3. Enter the key in the app

## Usage

### Basic Conversion

1. Click "Open Text File" or paste text into the editor
2. Select TTS engine and voice
3. Adjust rate and volume as needed
4. Click "Convert to Audio"
5. Select output folder

### SSML Tags

Use SSML tags for advanced control:
```xml
<emphasis level="strong">Important text</emphasis>
<break time="1s"/>
<prosody rate="slow">Slow speech</prosody>
<say-as interpret-as="date">12/25/2024</say-as>