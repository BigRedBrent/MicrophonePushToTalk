#include-once
#include <WinAPI.au3>
#include <WinAPIEx.au3>
#include <WinAPIMisc.au3>

Func _Beep($iFrequency = 500, $imsDuration = 1000, $iVolume = 100)

    Local Const $PI = 3.141592653589
    Local Const $TAU = 2 * $PI
    Local Const $iFormatChunkSize = 16
    Local Const $iHeaderSize = 8
    Local Const $iFormatType = 1
    Local Const $iTracks = 1
    Local Const $iSamplesPerSecond = 44100
    Local Const $iBitsPerSample = 16
    Local Const $iFrameSize = Floor($iTracks * (($iBitsPerSample + 7) / 8))
    Local Const $iBytesPerSecond = $iSamplesPerSecond * $iFrameSize
    Local Const $iWaveSize = 4
    Local Const $iSamples = Floor($iSamplesPerSecond * $imsDuration / 1000)
    Local Const $iDataChunkSize = Int($iSamples * $iFrameSize)
    Local Const $iFileSize = ($iWaveSize + $iHeaderSize + $iFormatChunkSize + $iHeaderSize + $iDataChunkSize)
    Local Const $sTagWAVHeader = "char ChunkID[4];int ChunkSize;char Format[4];char Subchunk1ID[4];" & _
            "int Subchunk1Size;short AudioFormat;short NumChannels;int SampleRate;" & _
            "int ByteRate;short BlockAlign;short BitsPerSample;char Subchunk2ID[4];" & _
            "int Subchunk2Size;"

    Local $sTagWAVHeaderAndData = $sTagWAVHeader & "short AudioWave[" & $iDataChunkSize / 2 & "]"
    Local $tWAVHeader_Data = DllStructCreate($sTagWAVHeaderAndData)

    $tWAVHeader_Data.ChunkID = "RIFF"
    $tWAVHeader_Data.ChunkSize = $iFileSize
    $tWAVHeader_Data.Format = "WAVE"
    $tWAVHeader_Data.Subchunk1ID = "fmt "
    $tWAVHeader_Data.Subchunk1Size = $iFormatChunkSize
    $tWAVHeader_Data.AudioFormat = $iFormatType
    $tWAVHeader_Data.NumChannels = $iTracks
    $tWAVHeader_Data.SampleRate = $iSamplesPerSecond
    $tWAVHeader_Data.ByteRate = $iBytesPerSecond;
    $tWAVHeader_Data.BlockAlign = $iFrameSize
    $tWAVHeader_Data.BitsPerSample = $iBitsPerSample
    $tWAVHeader_Data.Subchunk2ID = "data"
    $tWAVHeader_Data.Subchunk2Size = $iDataChunkSize

    Local $iAmplitude = 16383 / (100 / $iVolume)
    Local $iShort = 0
    $iTheta = $iFrequency * $TAU / $iSamplesPerSecond
    For $iStep = 0 To $iSamples - 1
        $iShort = Int($iAmplitude * Sin($iTheta * ($iStep)))
        DllStructSetData($tWAVHeader_Data, "AudioWave", $iShort, $iStep + 1)
    Next

    _WinAPI_PlaySound(DllStructGetPtr($tWAVHeader_Data), $SND_MEMORY)

EndFunc