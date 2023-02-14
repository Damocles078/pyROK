$shortcut_path = $args[0]
$pythonw_path = $args[1]
$pyRoK_path = $args[2]
$shortcut_working_directory = $args[3]
$shortcut_icon_path = $args[4]


$sh = New-Object -comObject WScript.Shell
$shortcut = $sh.CreateShortcut($shortcut_path)
$shortcut.TargetPath = $pythonw_path
$shortcut.Arguments = $pyRoK_path
$shortcut.WorkingDirectory = $shortcut_working_directory
$shortcut.IconLocation = "$shortcut_icon_path, 0"
$shortcut.Description = "RoK python analysis tool"
$shortcut.Save()