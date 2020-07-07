# Windows 10 Text-to-Speech converter

# This file is used along with the ppt to tutorial creator to
# generate video tutorials from ppt

# Voice types: (not working as of now - default voice to be used)
# Microsoft Mark Mobile, Microsoft David Mobile, Microsoft Zira Mobile, Microsoft Eva Mobile

function Text2Audio([string]$TextFile, [string]$WavFile) {
	#$TextFile=$args[0]
	#$WavFile=$args[1]
	Write-Output "    > text2audio.ps1: Audio creation ... $TextFile -> $WavFile"
	$Text = [IO.File]::ReadAllText($TextFile)

	Add-Type -AssemblyName System.Speech
	$SpeechSynthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
	#$SpeechSynthesizer.SelectVoice("Microsoft Mark Mobile")
	$SpeechSynthesizer.Rate = 0  # -10 is slowest, 10 is fastest
	$SpeechSynthesizer.SetOutputToWaveFile($WavFile)

	$SpeechSynthesizer.Speak($Text)
	$SpeechSynthesizer.Dispose()
}

function CreateAudiofile([string]$TextFile, [string]$WavFile) {
	Write-Output "$TextFile -> $WavFile"
}

function CreateAudioFiles([string]$InFile) {
	$NotesFiles = Get-Content -Path $InFile
	
	$file_list = $NotesFiles.Split()
	
	foreach ($notesfile in $file_list) {
		$audio_file = "$($notesfile).wav"
		Text2Audio $notesfile $audio_file
	}
}

$InFile=$args[0]

CreateAudioFiles $InFile
