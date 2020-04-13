Imports System.Console
Imports System.Timers
'Imports WMPLib

' ##############################################################
' # MinVaders game by L.Minett (c) 2020 - based on Casio game  #
' # ---------------------------------------------------------- #
' # Ideas for improvement:                                     #
' # 1) select music at random                                  #
' # 2) Add difficulty:                                         #
' #  - could reduce number Of values used.e.g. 1-3             #
' #  - could be easy starts @2s per num                        #
' #  - could shorten or lengthen length of num string          #
' # 3) More complex leveling requirements and scoring          #
' #  - could be easy starts @2s per num                        #
' # 4) different music as numbers increase or lives decrease   #
' # 5) Add high score table                                    #
' # 6) Persist game settings to a file and allow user update   # 
' # 7) Mode to emulate original game (only aim and fire keys)  #
' # 8) Orig. gameplay where numbers worth more as go to the    # 
' #    left, any num can be cancelled, but those not scroll    # 
' #    across still, so game ends when any number reaches base # 
' #    leaving a strategy to wipe least sig. dig for as long   # 
' #    possible until others nums get too close to base        # 
' ##############################################################
Module module1
    ' game parameters
#Region "Parameters"
    ' these are constants as they are used to reset and should not be changed in-game.  They can be swapped for user parameters and stored in resources object
    Const UFOSymbol As Char = "n" ' the symbol representing UFP
    Const NumLen As Int16 = 10 ' max number of digits before life lost
    Const NumLives As Int16 = 3
    Const StartingSpeed As Int16 = 1400 ' number of ms until next number is placed
    Const NumReqPerRound As Int16 = 10 ' number of matches needed per round for level 1
    Const PointsPerMatch As Int16 = 20 ' number of points per match
    Const PointsMisFire As Int16 = 10 ' points deducted for misfiring
    Const ClearBonus As Int16 = 50 ' number of points given when clear number field
    Const UFOPoints As Int16 = 20 ' points given for eliminating UFO ('n')
    Const LevelUpBonus As Int16 = 50 ' points given when levelling up
#End Region
#Region "Declarations"
    ' Music/SFX pointers
    Dim PlaySound As Boolean = True ' Wil play sounds if true
    'Dim BackgroundMusic As New WindowsMediaPlayer
    Enum SoundEffect ' used when requesting a sound effect to be played
        ClearOne = 0
        ClearAll = 1
        Misfire = 2
        GameOver = 3
        LifeLost = 4
        LevelUp = 5
    End Enum
    Enum Music ' used when requesting bkg music to play
        Main = 0
        Intro = 1
    End Enum
    'Dim MusicPath As String = "D:\ICT & CS Dropbox\Lee Minett\Teaching Data\VB & VBA Code\Casio Invaders Game\My Project\" ' path is appended to all file names 

    ' make some variables global, so accessible by separate thread
    'Dim InputThreading As New Threading.Thread(AddressOf ThreadUserInput) ' initialise ThreadUserInput sub as a separate process ##
    Dim CurrValue As Int16 = 0 ' user's number
    Dim NumbersGenerated As Int16 = 0 ' tracks how many numbers have been generated (prevents more numbers than needed for level)
    Dim NumbersRemoved As Int16 = 0 ' tracks how many successful numbers have been removed (used for levelling)
    Dim NumPerRound As Int16 = NumReqPerRound ' number of matches needed per round
    Dim Lives As Int16 = NumLives ' tracks current lives
    Dim Level As Int16 = 0
    Dim Score As Integer = 0 ' player's score
    Dim NumString As String = "" ' these store the random generated values the user needs to clear
    Dim StateChange As Boolean = True ' set to true when the game interface needs to be updated. Used to prevent re-drawing game too often
    Dim MSUntilNextNumber As Int16 = StartingSpeed ' this will decrease as the levels go up
    Dim Watch As New Stopwatch() ' used to time when new numbers appear 

#End Region

    Sub Main()
        ' Casio Invaders Game

        Console.CursorVisible = False ' prevent cursor being drawn on the screen

        StartScreen() ' load intro screen

        ' set/reset up game
        ResetGame()

        ' generate first number
        AddNumber()

        ' start music - until find concurrent audio solution, leave out background music
        'PlayMusic(Music.Main)

        'InputThreading.Start() ' start running thread
        Watch.Start() ' start timer for next number
        StateChange = True ' ensure screen redrawn

        Do While Lives > 0
            Do Until StateChange
                ' this will keep looping until some aspect makes a change to the game's state
                ' this prevents the console being drawn more often than needed
                ' while looping, check to see if another number should be shown

                If Console.KeyAvailable Then ' has the user pressed a key, if so, read it
                    Select Case Console.ReadKey(True).Key ' do not echo the key press to the screen
                    ' note: -1 demotes 'n' in the game
                        Case ConsoleKey.UpArrow
                            CurrValue = If(CurrValue = 9, -1, CurrValue + 1)
                            StateChange = True
                        Case ConsoleKey.DownArrow
                            CurrValue = If(CurrValue = -1, 9, CurrValue - 1)
                            StateChange = True
                        Case ConsoleKey.Spacebar
                            Fire()
                            StateChange = True
                        Case ConsoleKey.Q
                            End
                    End Select
                End If
                If Watch.ElapsedMilliseconds >= MSUntilNextNumber Then
                    ' time to insert a new number
                    AddNumber()
                    StateChange = True ' break out of loop to display and caculate game values
                    Watch.Restart() ' start watch over
                End If
            Loop
            StateChange = False ' reset back
            ' check if lives have gone
            If NumString.Length > NumLen Then ' lose life
                LoseLife()
            ElseIf NumbersRemoved >= NumPerRound Then ' level up
                LevelUp()
            End If
            DisplayGame()

        Loop

        ' ToDo: move this out to separate sub, with gfx, high score, etc 

        Console.WriteLine("GAME OVER")

        'InputThreading.Suspend() ' stop responding to user input
        Console.ReadKey()
    End Sub
    Sub AddNumber()
        ' simply adds a random number to the number string
        ' if last number of the round, insert a space
        If NumbersGenerated >= NumPerRound Then ' once reach num limit per round, insert space
            NumString &= " "
        Else
            Dim NextNum As New System.Random() ' will allow generation of random number
            Randomize()

            If NextNum.Next(0, 20) = 7 Then ' arbitrary value (1 in 20 chance of seeing UFO)
                NumString &= UFOSymbol
            Else ' normal number
                NumString &= NextNum.Next(0, 10).ToString ' picks a random starting number between 0 and 9
            End If

            NumbersGenerated += 1 ' increment one number  
        End If
    End Sub
    Sub LoseLife()
        ' called when number string goes over length (i.e. len > numlen)
        ' separated for easy enhancements, modifications, etc

        Lives -= 1
        If Lives = 0 Then
            PlaySFX(SoundEffect.GameOver)
        Else
            PlaySFX(SoundEffect.LifeLost)
            NumbersRemoved = 0 ' reset level 
            NumbersGenerated = 0
            NumString = ""

            AddNumber() ' add first number
        End If

    End Sub
    Sub LevelUp()
        ' Runs when user levels up

        PlaySFX(SoundEffect.LevelUp)
        NumbersRemoved = 0 ' reset number got
        NumbersGenerated = 0 ' reset
        Level += 1
        Score += LevelUpBonus
        NumString = "" ' clear number string

        ' increase number speed by 10%
        MSUntilNextNumber = Math.Ceiling(MSUntilNextNumber * 0.9)
        NumPerRound = Math.Ceiling(NumPerRound * 1.2) ' increase number of matches needed for next level by 20%

        ' ToDo: what else needs to happen when level up? 


    End Sub
    Sub ResetGame()
        ' sets game values to default for next game
        CurrValue = 0 ' user's starting value
        Score = 0 ' player score
        Level = 1 ' player round
        Lives = NumLives ' reset to max num lives
        MSUntilNextNumber = StartingSpeed ' # muliseconds until next number is shown
        NumbersRemoved = 0 ' no numbers removed
        NumPerRound = NumReqPerRound ' number of matches needed per round
        NumbersGenerated = 0 ' reset numbers generated

    End Sub
    Sub Fire()
        ' handles user firing value
        ' deals with removing values from string and calculating points

        Dim UserValue As String = If(CurrValue = -1, UFOSymbol, CurrValue.ToString)

        ' ToDo: change score to be proportion based on position?

        ' is current value same as left most number (first check there's something in the string and it's numeric)
        If NumString.Trim.Length > 0 AndAlso UserValue = Left(NumString, 1) Then
            ' value matches current value

            If NumString.Length = 1 Then ' only a single char in numstring
                ' award point + bonus
                PlaySFX(SoundEffect.ClearAll)
                If UserValue = UFOSymbol Then ' either add points for UFO or regular symbol
                    Score += UFOPoints + ClearBonus
                Else
                    Score += PointsPerMatch + ClearBonus
                End If
                NumString = ""
            Else
                PlaySFX(SoundEffect.ClearOne)
                NumString = Mid(NumString, 2) ' remove left most char
                If UserValue = UFOSymbol Then ' either add points for UFO or regular symbol
                    Score += UFOPoints
                Else
                    Score += PointsPerMatch
                End If
            End If
            NumbersRemoved += 1 ' increase number removed
        Else ' misfire
            PlaySFX(SoundEffect.Misfire)
            Score -= PointsMisFire ' deduct penalty for misfire
        End If

    End Sub
    Sub StartScreen()
        Clear() ' clear the screen

        PlayMusic(Music.Intro)
        WriteLine("Welcome to MinVaders." & vbCrLf)
        WriteLine("Keys:")
        WriteLine("    [up]    - Increase number")
        WriteLine("    [down]  - Decrease number")
        WriteLine("    [space] - Fire")
        WriteLine("    Q       - Quit" & vbCrLf)
        WriteLine("Press any key to start")

        Console.ReadKey(True)
        StopMusic()
        Console.Clear()
    End Sub
    Sub DisplayGame()
        ' this sub updates the gameplay screen 
        ' data is overwritten, not cleared so as to maintain non-flicker gameplay

        ' display HUD

        ForegroundColor = If(Score >= 0, ConsoleColor.Blue, ConsoleColor.Red)
        SetCursorPosition(1, 0)
        WriteLine("Score:" & Score.ToString.PadLeft(6, " ")) ' build score string - keep right hand of score in same position)

        ' display lives
        ForegroundColor = ConsoleColor.Cyan
        SetCursorPosition(18, 0)
        WriteLine("Lives:" & Lives.ToString.PadLeft(3, " ")) ' show number of lives

        ' display level
        ForegroundColor = ConsoleColor.DarkCyan
        SetCursorPosition(34, 0)
        WriteLine("Level:" & Level.ToString.PadLeft(3, " ")) ' show current level

        ' display num numbers needed for next level
        ForegroundColor = ConsoleColor.Gray
        SetCursorPosition(48, 0)
        WriteLine("Num. matches req.:" & (NumPerRound - NumbersRemoved).ToString.PadLeft(3, " ")) ' how many numbers until next level

        ' display playfield
        ForegroundColor = ConsoleColor.White
        SetCursorPosition(1, 3)
        WriteLine(If(CurrValue = -1, UFOSymbol, CurrValue.ToString)) ' show player's current number or the ufo symbol

        ForegroundColor = ConsoleColor.Green
        SetCursorPosition(3, 3)
        WriteLine("=") ' show barrier

        ' display number string in green for under half-fill
        ' display number string in amber for half-filled
        ' display number string in red for 2 from max len
        Dim NumStrClr As New ConsoleColor
        NumStrClr = If(NumString.Length < NumLen / 2, ConsoleColor.Green, If(NumString.Length > NumLen - 2, ConsoleColor.Red, ConsoleColor.Yellow))
        ForegroundColor = NumStrClr
        SetCursorPosition(4, 3)
        WriteLine(NumString.PadLeft(NumLen + 1, " "))


    End Sub

    Sub PlaySFX(SF As SoundEffect)
        ' play corresponding track
        If PlaySound Then ' 
            Try ' handle well if error playing
                If SF = SoundEffect.ClearAll Then
                    My.Computer.Audio.Play(My.Resources.Harpsichord, AudioPlayMode.Background)
                ElseIf SF = SoundEffect.ClearOne Then
                    My.Computer.Audio.Play(My.Resources.WHOP, AudioPlayMode.Background)
                ElseIf SF = SoundEffect.GameOver Then
                    ' pick track at random
                    Dim Tr As New Random
                    Randomize()
                    Dim Ran As Int16 = Tr.Next(0, 2)
                    If Ran = 0 Then
                        My.Computer.Audio.Play(My.Resources.Not_Afraid_To_Run___Jury_6, AudioPlayMode.Background)
                    ElseIf Ran = 1 Then
                        My.Computer.Audio.Play(My.Resources.Cristal_Sky___Vocal_Accents, AudioPlayMode.Background)
                    End If
                ElseIf SF = SoundEffect.LifeLost Then
                    My.Computer.Audio.Play(My.Resources.TWANG, AudioPlayMode.Background)
                ElseIf SF = SoundEffect.Misfire Then
                    My.Computer.Audio.Play(My.Resources.Thud, AudioPlayMode.Background)
                ElseIf SF = SoundEffect.LevelUp Then
                    ' pick track at random
                    Dim Tr As New Random
                    Randomize()
                    Dim Ran As Int16 = Tr.Next(0, 3)
                    If Ran = 0 Then
                        My.Computer.Audio.Play(My.Resources.Brute___Power_Failure__5s_, AudioPlayMode.Background) ' could change to wait until finished
                    ElseIf Ran = 1 Then
                        My.Computer.Audio.Play(My.Resources.Cristal_Sky___Vocal_Accents, AudioPlayMode.Background) ' could change to wait until finished
                    ElseIf Ran = 2 Then
                        My.Computer.Audio.Play(My.Resources.Step_Forward___Strummer, AudioPlayMode.Background) ' could change to wait until finished
                    End If
                End If
            Catch ex As Exception
                ' track error, ignore
            End Try
        End If
    End Sub
    Sub PlayMusic(Track As Music)

        'ToDo: Concurrent music using Media player: 
        ' https://social.msdn.microsoft.com/Forums/vstudio/en-US/b47d3599-4263-4429-b62f-86145829299d/background-music-while-sound-fx-plays?forum=vbgeneral
        ' https://docs.microsoft.com/en-us/windows/win32/wmp/creating-the-windows-media-player-control-programmatically
        ' did not work when attempted


        ' handles requests to play music

        StopMusic()

        If PlaySound Then
            Try ' handle well if error playing
                If Track = Music.Main Then
                    'BackgroundMusic.URL = MusicPath & "Synthetic Dream - Atmospheric.wav"
                    'BackgroundMusic.controls.play()
                    My.Computer.Audio.Play(My.Resources.Synthetic_Dream___Atmospheric, AudioPlayMode.Background)
                ElseIf Track = Music.Intro Then
                    My.Computer.Audio.Play(My.Resources.Always_Moving_Forward__2m_, AudioPlayMode.Background)
                End If
            Catch ex As Exception
                ' track error, ignore
            End Try
        End If
    End Sub
    Sub StopMusic()
        ' stops all background sound
        My.Computer.Audio.Stop()
    End Sub
End Module