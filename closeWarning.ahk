WinActivate, Microsoft Excel
While !WinExist("Excel Compatibility Notice")
{
    Sleep, 200
}
WinClose Excel Compatibility Notice