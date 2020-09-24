Attribute VB_Name = "AktScreen"
Function AktAufloesung$()
  TPX = Screen.TwipsPerPixelX
  TPY = Screen.TwipsPerPixelY
  Wdth = Screen.Width / TPX
  Hght = Screen.Height / TPY
  AktAufloesung$ = Wdth & " x " & Hght
End Function

