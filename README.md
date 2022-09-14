# kaixi

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$signature=@'
[DllImport("user32.dll",CharSet=CharSet.Auto,CallingConvention=CallingConvention.StdCall)]
public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@
$SendMouseClick = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru


$i = 0
$X = 500
$Y = 50

while ($i -lt 3) {
    Write-Output "Start Operation"

    Start-Sleep -s 1

    # マウスカーソル移動

    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)

    Start-Sleep -s 1

    # 左クリック
    #Write-Output "左クリック"
    #$SendMouseClick::mouse_event(0x0008, 0, 0, 0, 0);
    #$SendMouseClick::mouse_event(0x0010, 0, 0, 0, 0);
    # 右クリック
    #Write-Output "右クリック"
    #$SendMouseClick::mouse_event(0x0002, 0, 0, 0, 0);
    #$SendMouseClick::mouse_event(0x0004, 0, 0, 0, 0);

    Write-Output "End Operation"
    $X = $X - 1

}

Read-Host
