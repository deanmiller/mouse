Function J
{

  <#
  .Synopsis
   Function to keep Windows Session active

  .Description
   This script is technicaly not a mouse jiggler. It is called that because most people are familar with this term. 
   This script presses the "Ctrl" button every 20 to 30 seconds randomly which will keep your computer active.

  .Example
   Start-MouseJiggler

  .Example
   # You can set this script to turn off in any amount of minutes, for example if you want an extended lunch. 
   # Your computer will go to sleep after the script stops at your choosen time.
   
   J -Minutes 30

  .Link
   about_functions
   about_functions_advanced
   about_functions_advanced_methods
   about_functions_advanced_parameters

  .Notes
   NAME:      Start-MouseJiggler
   AUTHOR:    Dean Miller
   LASTEDIT:  6/1/2021
   #Requires -Version 3.0
  #>

  [CmdletBinding()]
  param(
  [Parameter(Position=0)]
  [Alias('Hours')]
  $RunForXHours,
  [Parameter(Position=1)]
  [Alias('Minutes')]
  $RunForXMinutes,
  [Parameter(Position=2)]
  [switch]$NoLunchSleep = $false
  )


  Begin {

    #region --- Hide/Show PowerShell Window ---
      
      #https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/show-or-hide-windows
    
      Enum ShowStates
      {
        Hide = 0
        Normal = 1
        Minimized = 2
        Maximized = 3
        ShowNoActivateRecentPosition = 4
        Show = 5
        MinimizeActivateNext = 6
        MinimizeNoActivate = 7
        ShowNoActivate = 8
        Restore = 9
        ShowDefault = 10
        ForceMinimize = 11
      }
      
      $Code = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
      $Type = Add-Type -MemberDefinition $Code -Name myAPI -PassThru
      $hwnd = (Get-Process -Id $PID).MainWindowHandle
    
    #endregion --------------------------------
  
  }


  Process
  {
   
    $TitleSnapshot = $host.ui.RawUI.WindowTitle

    function Trap-CtrlC {
      ## Stops Ctrl+C from exiting this function
      [console]::TreatControlCAsInput = $true
      ## And you have to check every keystroke to see if it's a Ctrl+C
      ## As far as I can tell, for CtrlC the VirtualKeyCode will be 67 and
      ## The ControlKeyState will include either "LeftCtrlPressed" or "RightCtrlPressed"
      ## Either of which will -match "CtrlPressed"
      ## But the simplest thing to do is just compare Character = [char]3
      if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character))
      {
        $host.ui.RawUI.WindowTitle = $TitleSnapshot
        cls
        break
      }
    }
    
    if($RunForXHours -eq $null -and $RunForXMinutes -eq $null){
      $RunForXHours = 9
      $RunForXMinutes = 0
    }
    
    $StopTime = (Get-Date).AddHours($RunForXHours).AddMinutes($RunForXMinutes)
    
    <#
     This script stops no later than hour 17 which is 5pm
     You don't want this to run all evening. 
     That would be a red flag to most employers that you are using a Mouse Jiggler.
     And that you might not be working during every second of the day.
    #>
    $NoLaterThan = '17'
    
    #region --- Format stop time for screen display ---
    
      if($StopTime.Hour -ge $NoLaterThan -or $StopTime.Hour -le 4){
      
        $UnformattedTime = Get-Date -Hour $NoLaterThan -Minute 0 | Get-Date -Format g
        
        # The .Substring($UnformattedTime.IndexOf(' ')) method will find the first space in a string & start from there.
        $Time = $UnformattedTime.Substring($UnformattedTime.IndexOf(' ')).trim()
      
      }else{
      
        $UnformattedTime = $StopTime | Get-Date -Format g
        $Time = $UnformattedTime.Substring($UnformattedTime.IndexOf(' ')).trim()
      
      }
    
    #endregion ----------------------------------------
    
    Write-Host "`nStop time is $Time." -ForegroundColor Cyan

    if($NoLunchSleep -eq $false){
    
      Write-Host "`nThe mouse jiggler will also stop at noon for lunch.`nTo disable this use the" -ForegroundColor Cyan -NoNewline
      Write-Host " -NoLunchSleep" -ForegroundColor Yellow -NoNewline
      Write-Host " Parameter." -ForegroundColor Cyan
    
    }
    
    $host.ui.RawUI.WindowTitle = "$($TitleSnapshot.Split(':')[0]): $Time"
    
    1..3 | ForEach-Object {
      Trap-CtrlC  
      Start-Sleep -Seconds 1
    }

    $type::ShowWindowAsync($hwnd, [ShowStates]::Minimized) | Out-Null

    cls
    
    $WShell = New-Object -ComObject wscript.shell;
    
    # while($true) = Loop forever.
    # The loop will only be stopped when the "break" keyword is executed.
    while($true){

      $LunchSleep = $false
      if($RunForXHours -gt 0 -and $((Get-Date).Hour) -eq 12 -and $NoLunchSleep -eq $false){$LunchSleep = $true}
      
      if(!$LunchSleep){$WShell.SendKeys('^')} # ^ = Ctrl

      $SleepNumber = 20..30 | Get-Random 
      1..$SleepNumber | ForEach-Object {
        Trap-CtrlC  
        Start-Sleep -Seconds 1
      } # Sleep for 20 to 30 seconds

      if((Get-Date) -ge $StopTime -or $((Get-Date).Hour) -ge $NoLaterThan){
        $host.ui.RawUI.WindowTitle = $TitleSnapshot
        Get-Date
        $type::ShowWindowAsync($hwnd, [ShowStates]::Restore) | Out-Null
        break
      }
    
    }

  }# End Process

}# End Function J
