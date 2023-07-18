# Clear variables for repeatability
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

Function Invoke-Note {
    param (
        $Note,
        $Duration
    )
    # If ForceMaxVolume is set to True...
    If ($script:ForceMaxVolume -eq $true) {
        # Use SendKeys to increment max volume as specified in parent function
        # (smaller values to avoid timing delays)
        1..10 | % {$script:wshShell.SendKeys([char]175)}
    }
    # Determine if note is a rest
    If ($Note -eq 'REST') {
        # Pause for specified duration
        Start-Sleep -Milliseconds $Duration
    } Else {
        # Generate tone for note/duration
        [console]::Beep($Note, $Duration)
    }
} # End Function Invoke-Note

Function Invoke-Tune {
    param (
        [bool]$ForceMaxVolume
    )
    # If ForceMaxVolume is set to True...
    If ($ForceMaxVolume -eq $true) {
        # set variable w/ script scope
        $script:ForceMaxVolume = $true
        # Use SendKeys to max volume level
        $wshShell = new-object -com wscript.shell;1..50 | % {$wshShell.SendKeys([char]175)}
        # Prime wshShell for incrementation
        $script:wshShell = new-object -com wscript.shell
    }
    # Set Beats Per Minute
    $BPM = 120
    # Correspond tone frequencies & notes
    $Note = New-Object PSObject -Property @{
        REST = 'REST'
        DS1  =  38.89
        E1   =  41.20
        F1   =  43.65
        FS1  =  46.25
        G1   =  49.00
        GS1  =  51.91
        A1   =  55.00
        AS1  =  58.27
        B1   =  61.74
        C2   =  65.41
        CS2  =  69.30
        D2   =  73.42
        DS2  =  77.78
        E2   =  82.41
        F2   =  87.31
        FS2  =  92.50
        G2   =  98.00
        GS2  =  103.83
        A2   =  110.00
        AS2  =  116.54
        B2   =  123.47
        C3   =  130.81
        CS3  =  138.59
        D3   =  146.83
        DS3  =  155.56
        E3   =  164.81
        F3   =  174.61
        FS3  =  185.00
        G3   =  196.00
        GS3  =  207.65
        A3   =  220.00
        AS3  =  233.08
        B3   =  246.94
        C4   =  261.63
        CS4  =  277.18
        D4   =  293.66
        DS4  =  311.13
        E4   =  329.63
        F4   =  349.23
        FS4  =  369.99
        G4   =  392.00
        GS4  =  415.30
        A4   =  440.00
        AS4  =  466.16
        B4   =  493.88
        C5   =  523.25
        CS5  =  554.37
        D5   =  587.33
        DS5  =  622.25
        E5   =  659.26
        F5   =  698.46
        FS5  =  739.99
        G5   =  783.99
        GS5  =  830.61
        A5   =  880.00
        AS5  =  932.33
        B5   =  987.77
        C6   =  1046.50
        CS6  =  1108.73
        D6   =  1174.66
        DS6  =  1244.51
        E6   =  1318.51
        F6   =  1396.91
        FS6  =  1479.98
        G6   =  1567.98
        GS6  =  1661.22
        A6   =  1760.00
        AS6  =  1864.66
        B6   =  1975.53
        C7   =  2093.00
        CS7  =  2217.46
        D7   =  2349.32
        DS7  =  2489.02
        E7   =  2637.02
        F7   =  2793.83
        FS7  =  2959.96
        G7   =  3135.96
        GS7  =  3322.44
        A7   =  3520.00
        AS7  =  3729.31
        B7   =  3951.07
        C8   =  4186.01
        CS8  =  4434.92
        D8   =  4698.64
        DS8  =  4978.03
    }    
    # Define duration of notes for BPM
    $Duration = New-Object PSObject -Property @{
        Whole =           ((60000/$BPM) * 4)
        DottedHalf =      ((60000/$BPM) * 3.5)
        Half =            ((60000/$BPM) * 2)
        DottedQuarter =   ((60000/$BPM) * 1.5)
        Quarter =          (60000/$BPM)
        DottedEighth =    ((60000/$BPM) * 0.75)
        Eighth =          ((60000/$BPM) * 0.5)
        DottedSixteenth = ((60000/$BPM) * 0.375)
        Sixteenth =       ((60000/$BPM) * 0.25)
    }
    # INSERT NOTES BELOW #
    
    # INSERT NOTES ABOVE #
} # End Function Invoke-Tune

# Use -ForceMaxVolume boolean param to force maximum playback volume
# For best performance, do not force maximum playback volume

Invoke-Tune -ForceMaxVolume $True
