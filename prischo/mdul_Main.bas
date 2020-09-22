Attribute VB_Name = "mdul_Main"
Option Explicit

Public pIndex As Integer
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Function ReleaseCapture Lib "user32" () As Long

Sub Main()
frmlog.Show
End Sub

Public Function DragForm(frm As Form)
  Dim ret As Long
  ret = ReleaseCapture()
  ret = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, 2&, 0&)
End Function
Public Sub form1init()
With frmaddstud
.ctrl_SkinableForm.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_SkinableForm.BackColor = &H0&
        .ctrl_SkinableForm.CaptionTop = 240
        .ctrl_SkinableForm.CaptionColor = &H0&
      '  .ctrl_ListObject1.SkinPath = App.Path & "\Skins\TreasureChest"
      '  .ctrl_ListObject1.ForeColor = &H0&
       ' .ctrl_ListObject1.MouseMoveColor = &H0&
        '.ctrl_ListObject1.MouseDownColor = &H0&
       '.ctrl_ListObject1.DrawMenu
        
        ' .ctrl_ListObject.SkinPath = App.Path & "\Skins\TreasureChest"
        '.ctrl_ListObject.ForeColor = &H0&
        '.ctrl_ListObject.MouseMoveColor = &H0&
        '.ctrl_ListObject.MouseDownColor = &H0&
       '.ctrl_ListObject.DrawMenu
        Call frmaddstud.ctrl_SkinableForm.LoadSkin(Form1)
        'Call Form1.ctrl_ListObject.AddItem("Register")
        'Call Form1.ctrl_ListObject.AddItem("Check account Balance")
        'Call Form1.ctrl_ListObject.AddItem("My Account")
        'Call Form1.ctrl_ListObject.AddItem("Log Out")
      '  .ctrl_SkinableButton.SkinPath = App.Path & "\Skins\TreasureChest"
      '  .ctrl_SkinableButton.ForeColor = &HFFFFFF
      ' .ctrl_SkinableButton.LoadSkin
       ' .ctrl_SkinableButton.Refresh
     '  .ctrl_PullDownMenu.AddItem ("File")
       
    '.ctrl_PullDownMenu.BackColor = &H0&
     '   .ctrl_PullDownMenu.ForeColor = &HFFFFFF
      '  .ctrl_PullDownMenu.Refresh
End With
End Sub


Public Sub ChangeSkinToDefault()
    With frmmain
     '   .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Default"
       ' .ctrl_SkinableForm.BackColor = &HCECECE
      '  .ctrl_SkinableForm.CaptionTop = 360
        '.ctrl_SkinableForm.CaptionColor = &H0&
        'Call frmmain.ctrl_SkinableForm.LoadSkin(frmmain)
        
        '.ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Default"
        '.ctrl_btn_Previous.ForeColor = &H0&
        '.ctrl_btn_Previous.LoadSkin
        '.ctrl_btn_Previous.Refresh
        '.ctrl_btn_Next.SkinPath = App.Path & "\Skins\Default"
        '.ctrl_btn_Next.ForeColor = &H0&
        '.ctrl_btn_Next.LoadSkin
        '.ctrl_btn_Next.Refresh
        '.ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Default"
        '.ctrl_btn_Exit.ForeColor = &H0&
        '.ctrl_btn_Exit.LoadSkin
        '.ctrl_btn_Exit.Refresh
        
        '.ctrl_ListObject.SkinPath = App.Path & "\Skins\Default"
        '.ctrl_ListObject.ForeColor = &H0&
        '.ctrl_ListObject.MouseMoveColor = &H0&
        '.ctrl_ListObject.MouseDownColor = &HC0&
        '.iml_Toolbar.ListImages.Clear
        '.iml_Toolbar.ListImages.add 1, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Back.gif")
        '.iml_Toolbar.ListImages.add 2, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Forward.gif")
        '.iml_Toolbar.ListImages.add 3, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Home.gif")
        '.iml_Toolbar.ListImages.add 4, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Refresh.gif")
        '.iml_Toolbar.ListImages.add 5, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Open.gif")
        '.iml_Toolbar.ListImages.add 6, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Document.gif")
        '.iml_Toolbar.ListImages.add 7, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Search.gif")
        '.iml_Toolbar.ListImages.add 8, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Help.gif")
        '.iml_Toolbar.ListImages.add 9, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Stop.gif")
        '.ctrl_Toolbar.UnloadButtons
        '.ctrl_Toolbar.IconLeft = 60
        '.ctrl_Toolbar.IconTop = 60
        'Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        'Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        '.ctrl_ListObject.DrawMenu
        
        '.ctrl_Toolbar.SkinPath = App.Path & "\Skins\Default"
        '.ctrl_Toolbar.BackColor = &HCECECE
        '.ctrl_Toolbar.DrawToolbar
        '.ctrl_Toolbar.Refresh
        
        '.ctrl_Panel.SkinPath = App.Path & "\Skins\Default"
        '.ctrl_Panel.DrawPanel
        
        '.ctrl_PullDownMenu.BackColor = &HCECECE
        ''.ctrl_PullDownMenu.ForeColor = &H0&
        '.ctrl_PullDownMenu.Refresh
        
       ' .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Default"
       ' .ctrl_ChannelBar.SubItemTop = 370
        '.ctrl_ChannelBar.MouseMoveColor = &H0&
        '.ctrl_ChannelBar.MouseDownColor = &H0&
        '.ctrl_ChannelBar.SubMouseMoveColor = &H0&
        '.ctrl_ChannelBar.SubMouseDownColor = &H0&
        '.ctrl_ChannelBar.DrawMenu
        
       ' .Line1.BorderColor = &H0&
       ' .lbl_Statusbar.ForeColor = &H0&
        
       ' .pic_Viewport.BackColor = &H0&
        '.tbx_Text.BackColor = &H0&
        '.tbx_Text.ForeColor = &HFFFFFF
    End With
End Sub

Public Sub ChangeSkinToTitanium()
    With frmmain
       ' .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Titanium"
        '.ctrl_SkinableForm.BackColor = &H4B4A4A
        '.ctrl_SkinableForm.CaptionTop = 195
        '.ctrl_SkinableForm.CaptionColor = &H0&
        'Call frmmain.ctrl_SkinableForm.LoadSkin(frmmain)
       '
      '  .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Titanium"
      '  .ctrl_btn_Previous.ForeColor = &HFFFFFF
       ' .ctrl_btn_Previous.LoadSkin
        '.ctrl_btn_Previous.Refresh
        '.ctrl_btn_Next.SkinPath = App.Path & "\Skins\Titanium"
      '  .ctrl_btn_Next.ForeColor = &HFFFFFF
      '  .ctrl_btn_Next.LoadSkin
     '   .ctrl_btn_Next.Refresh
     '   .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Titanium"
     '   .ctrl_btn_Exit.ForeColor = &HFFFFFF
      '  .ctrl_btn_Exit.LoadSkin
      '  .ctrl_btn_Exit.Refresh
        
      '  .ctrl_ListObject.SkinPath = App.Path & "\Skins\Titanium"
      '  .ctrl_ListObject.ForeColor = &H0&
      '  .ctrl_ListObject.MouseMoveColor = &H0&
       ' .ctrl_ListObject.MouseDownColor = &H0&
       ' .iml_Toolbar.ListImages.Clear
       ' .iml_Toolbar.ListImages.add 1, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Back.gif")
       ' .iml_Toolbar.ListImages.add 2, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Forward.gif")
       ' .iml_Toolbar.ListImages.add 3, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Home.gif")
       '' .iml_Toolbar.ListImages.add 4, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Refresh.gif")
       ' .iml_Toolbar.ListImages.add 5, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Open.gif")
       ' .iml_Toolbar.ListImages.add 6, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Document.gif")
       ' .iml_Toolbar.ListImages.add 7, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Search.gif")
      '  .iml_Toolbar.ListImages.add 8, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Help.gif")
      '  .iml_Toolbar.ListImages.add 9, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Stop.gif")
      '  .ctrl_Toolbar.UnloadButtons
      '  .ctrl_Toolbar.IconLeft = 90
      '  .ctrl_Toolbar.IconTop = 90
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
       '' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        'Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
       '' Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
       ' .ctrl_ListObject.DrawMenu
        
       ' .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Titanium"
        ''.ctrl_Toolbar.BackColor = &H4B4A4A
        '.ctrl_Toolbar.DrawToolbar
        '.ctrl_Toolbar.Refresh
       '
       ' .ctrl_Panel.SkinPath = App.Path & "\Skins\Titanium"
       ' .ctrl_Panel.DrawPanel
       '
       ' .ctrl_PullDownMenu.BackColor = &H4B4A4A
       ' .ctrl_PullDownMenu.ForeColor = &HFFFFFF
       ' .ctrl_PullDownMenu.Refresh
       '
       ' .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Titanium"
       ' .ctrl_ChannelBar.SubItemTop = 395
       ' .ctrl_ChannelBar.MouseMoveColor = &H0&
       ' .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
       ' .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
       ' .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
       ' .ctrl_ChannelBar.DrawMenu
       '
       ' .Line1.BorderColor = &HFFFFFF
       ' .lbl_Statusbar.ForeColor = &HFFFFFF
       '
      '  .pic_Viewport.BackColor = &H0&
      '  .tbx_Text.BackColor = &H0&
     '   .tbx_Text.ForeColor = &HFFFFFF
    End With
End Sub

Public Sub ChangeSkinToBlue()
    With frmmain
     '   .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Blue"
      '  .ctrl_SkinableForm.BackColor = &HBD6E06
       ' .ctrl_SkinableForm.CaptionTop = 250
        '.ctrl_SkinableForm.CaptionColor = &H0&
        'Call frmmain.ctrl_SkinableForm.LoadSkin(frmmain)
        
       ' .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Blue"
       ' .ctrl_btn_Previous.ForeColor = &H0&
        '.ctrl_btn_Previous.LoadSkin
        '.ctrl_btn_Previous.Refresh
        '.ctrl_btn_Next.SkinPath = App.Path & "\Skins\Blue"
        '.ctrl_btn_Next.ForeColor = &H0&
        '.ctrl_btn_Next.LoadSkin
        '.ctrl_btn_Next.Refresh
        '.ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Blue"
        '.ctrl_btn_Exit.ForeColor = &H0&
        '.ctrl_btn_Exit.LoadSkin
        '.ctrl_btn_Exit.Refresh
        
        '.ctrl_ListObject.SkinPath = App.Path & "\Skins\Blue"
        '.ctrl_ListObject.ForeColor = &H0&
        '.ctrl_ListObject.MouseMoveColor = &H0&
        '.ctrl_ListObject.MouseDownColor = &H0&
       ' .iml_Toolbar.ListImages.Clear
       ' .iml_Toolbar.ListImages.add 1, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Back.gif")
       ' .iml_Toolbar.ListImages.add 2, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Forward.gif")
       ' .iml_Toolbar.ListImages.add 3, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Home.gif")
       ' .iml_Toolbar.ListImages.add 4, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Refresh.gif")
       ' .iml_Toolbar.ListImages.add 5, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Open.gif")
        '.iml_Toolbar.ListImages.add 6, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Document.gif")
        '.iml_Toolbar.ListImages.add 7, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Search.gif")
        ''.iml_Toolbar.ListImages.add 8, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Help.gif")
       ' .iml_Toolbar.ListImages.add 9, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Stop.gif")
       ' .ctrl_Toolbar.UnloadButtons
      '  .ctrl_Toolbar.IconLeft = 90
      '  .ctrl_Toolbar.IconTop = 90
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
      ''  Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
     '   Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
     '  ' Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
      ''  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        'Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        'Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        'Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        'Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        'Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        'Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
       ' .ctrl_ListObject.DrawMenu
        
      '  .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Blue"
       ' .ctrl_Toolbar.BackColor = &HBD6E06
       ' .ctrl_Toolbar.DrawToolbar
       ' .ctrl_Toolbar.Refresh
        
       ' .ctrl_Panel.SkinPath = App.Path & "\Skins\Blue"
       ' .ctrl_Panel.DrawPanel
        
       ' .ctrl_PullDownMenu.BackColor = &HBD6E06
       ' .ctrl_PullDownMenu.ForeColor = &H0&
       ' .ctrl_PullDownMenu.Refresh
        
      '  .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Blue"
      '  .ctrl_ChannelBar.SubItemTop = 440
      '  .ctrl_ChannelBar.MouseMoveColor = &H0&
      '  .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
      '  .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
      '  .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
      '  .ctrl_ChannelBar.DrawMenu
        
      '  .Line1.BorderColor = &H0&
      '  .lbl_Statusbar.ForeColor = &H0&
        
      '  .pic_Viewport.BackColor = &H571B02
       ' .tbx_Text.BackColor = &H571B02
      '  .tbx_Text.ForeColor = &HFFFFFF
    End With
End Sub

Public Sub ChangeSkinToDeco()
    With frmmain
      '  .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
      '  .ctrl_SkinableForm.BackColor = &HCECECE
      '  .ctrl_SkinableForm.CaptionTop = 300
      '  .ctrl_SkinableForm.CaptionColor = &H0&
      '  Call frm_Main.ctrl_SkinableForm.LoadSkin(frmmain)
      '
      '  .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Deco"
      '  .ctrl_btn_Previous.ForeColor = &H0&
      '  .ctrl_btn_Previous.LoadSkin
      '  .ctrl_btn_Previous.Refresh
      '  .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Deco"
      '  .ctrl_btn_Next.ForeColor = &H0&
      '  .ctrl_btn_Next.LoadSkin
      '  .ctrl_btn_Next.Refresh
      '  .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Deco"
     '   .ctrl_btn_Exit.ForeColor = &H0&
     '   .ctrl_btn_Exit.LoadSkin
      '  .ctrl_btn_Exit.Refresh
        
     '   .ctrl_ListObject.SkinPath = App.Path & "\Skins\Deco"
      '  .ctrl_ListObject.ForeColor = &H0&
     '   .ctrl_ListObject.MouseMoveColor = &H0&
     '   .ctrl_ListObject.MouseDownColor = &HC0&
     '   .iml_Toolbar.ListImages.Clear
     '   .iml_Toolbar.ListImages.add 1, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Back.gif")
     '   .iml_Toolbar.ListImages.add 2, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Forward.gif")
     '   .iml_Toolbar.ListImages.add 3, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Home.gif")
     '   .iml_Toolbar.ListImages.add 4, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Refresh.gif")
     '   .iml_Toolbar.ListImages.add 5, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Open.gif")
     '   .iml_Toolbar.ListImages.add 6, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Document.gif")
     '   .iml_Toolbar.ListImages.add 7, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Search.gif")
     '   .iml_Toolbar.ListImages.add 8, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Help.gif")
     '   .iml_Toolbar.ListImages.add 9, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Stop.gif")
     '   .ctrl_Toolbar.UnloadButtons
      '  .ctrl_Toolbar.IconLeft = 60
      '  .ctrl_Toolbar.IconTop = 60
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        'Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
       ' .ctrl_ListObject.DrawMenu
        
        '.ctrl_Toolbar.SkinPath = App.Path & "\Skins\Deco"
       ' .ctrl_Toolbar.BackColor = &HCECECE
       ' .ctrl_Toolbar.DrawToolbar
       ' .ctrl_Toolbar.Refresh
       '
       ' .ctrl_Panel.SkinPath = App.Path & "\Skins\Deco"
       ' .ctrl_Panel.DrawPanel
       '
       ' .ctrl_PullDownMenu.BackColor = &HCECECE
       ' .ctrl_PullDownMenu.ForeColor = &H0&
       ' .ctrl_PullDownMenu.Refresh
       '
       ' .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Deco"
       ' .ctrl_ChannelBar.SubItemTop = 400
       ' .ctrl_ChannelBar.MouseMoveColor = &H0&
       ' .ctrl_ChannelBar.MouseDownColor = &H0&
       ' .ctrl_ChannelBar.SubMouseMoveColor = &H0&
       ' .ctrl_ChannelBar.SubMouseDownColor = &H0&
       ' .ctrl_ChannelBar.DrawMenu
        
      '  .Line1.BorderColor = &H0&
       ' .lbl_Statusbar.ForeColor = &H0&
      '
      '  .pic_Viewport.BackColor = &H968A7B
      '  .tbx_Text.BackColor = &H968A7B
      '  .tbx_Text.ForeColor = &H0&
    End With
End Sub

Public Sub ChangeSkinToHolograph()
    With frmmain
      '  .ctr_SkinableForm.SkinPath = App.Path & "\Skins\Holograph"
      '  .ctrl_SkinableForm.BackColor = &H3A5959
      '  .ctrl_SkinableForm.CaptionTop = 285
      '  .ctrl_SkinableForm.CaptionColor = &HFFFFFF
      '  Call frmmain.ctrl_SkinableForm.LoadSkin(frmmain)
      '
      '  .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Holograph"
      '  .ctrl_btn_Previous.ForeColor = &HFFFFFF
      '  .ctrl_btn_Previous.LoadSkin
      '  .ctrl_btn_Previous.Refresh
      '  .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Holograph"
      ''  .ctrl_btn_Next.ForeColor = &HFFFFFF
     '   .ctrl_btn_Next.LoadSkin
     '   .ctrl_btn_Next.Refresh
     '   .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Holograph"
      '  .ctrl_btn_Exit.ForeColor = &HFFFFFF
       '' .ctrl_btn_Exit.LoadSkin
   '     .ctrl_btn_Exit.Refresh
    '
     '   .ctrl_ListObject.SkinPath = App.Path & "\Skins\Holograph"
      '  .ctrl_ListObject.MouseMoveColor = &HFFFFFF
  '     ' .ctrl_ListObject.MouseDownColor = &HFFFFFF
    '    .ctrl_ListObject.ForeColor = &HFFFFFF
   '     .iml_Toolbar.ListImages.Clear
    '    .iml_Toolbar.ListImages.add 1, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Back.gif")
 '       .iml_Toolbar.ListImages.add 2, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Forward.gif")
  '      .iml_Toolbar.ListImages.add 3, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Home.gif")
   '     .iml_Toolbar.ListImages.add 4, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Refresh.gif")
    '    .iml_Toolbar.ListImages.add 5, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Open.gif")
     '   .iml_Toolbar.ListImages.add 6, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Document.gif")
      '  .iml_Toolbar.ListImages.add 7, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Search.gif")
       ' .iml_Toolbar.ListImages.add 8, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Help.gif")
        '.iml_Toolbar.ListImages.add 9, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Stop.gif")
'        .ctrl_Toolbar.UnloadButtons
 '       .ctrl_Toolbar.IconLeft = 90
  '      .ctrl_Toolbar.IconTop = 90
   '     Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
    '    Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
     '   Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        'Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        'Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
 '       Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
  '      Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
   '     Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
    '    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
     '   Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
'       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
 '       Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
  '      Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
  '      Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
   '     .ctrl_ListObject.DrawMenu
    '
   '     .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Holograph"
    ''    .ctrl_Toolbar.BackColor = &H3A5959
       ' .ctrl_Toolbar.DrawToolbar
      '  .ctrl_Toolbar.Refresh
        
        '.ctrl_Panel.SkinPath = App.Path & "\Skins\Holograph"
  '      .ctrl_Panel.DrawPanel
   '
       ' .ctrl_PullDownMenu.BackColor = &H3A5959
    '    .ctrl_PullDownMenu.ForeColor = &HFFFFFF
     '   .ctrl_PullDownMenu.Refresh
      '
 '       .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Holograph"
  '      .ctrl_ChannelBar.SubItemTop = 395
   '     .ctrl_ChannelBar.MouseMoveColor = &H0&
    '    .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
     '   .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
      '  .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
       ' .ctrl_ChannelBar.DrawMenu
        '
'        .Line1.BorderColor = &HFFFFFF
 '       .lbl_Statusbar.ForeColor = &HFFFFFF
  '
   '     .pic_Viewport.BackColor = &H263C3C
    '    .tbx_Text.BackColor = &H263C3C
     '   .tbx_Text.ForeColor = &HFFFFFF
    End With
End Sub

Public Sub ChangeSkinToTreasureChest()
    With frmmain
'        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\TreasureChest"
 '       .ctrl_SkinableForm.BackColor = &H0&
  '      .ctrl_SkinableForm.CaptionTop = 240
   '     .ctrl_SkinableForm.CaptionColor = &H0&
     '   Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
    '
'        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\TreasureChest"
 '       .ctrl_btn_Previous.ForeColor = &HFFFFFF
  '      .ctrl_btn_Previous.LoadSkin
   '     .ctrl_btn_Previous.Refresh
     '   .ctrl_btn_Next.SkinPath = App.Path & "\Skins\TreasureChest"
    '    .ctrl_btn_Next.ForeColor = &HFFFFFF
     '   .ctrl_btn_Next.LoadSkin
      '  .ctrl_btn_Next.Refresh
'        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\TreasureChest"
 '       .ctrl_btn_Exit.ForeColor = &HFFFFFF
 '       .ctrl_btn_Exit.LoadSkin
  '      .ctrl_btn_Exit.Refresh
        
   '     .ctrl_ListObject.SkinPath = App.Path & "\Skins\TreasureChest"
    '    .ctrl_ListObject.ForeColor = &H0&
     '   .ctrl_ListObject.MouseMoveColor = &H0&
   '     .ctrl_ListObject.MouseDownColor = &H0&
      ''  .iml_Toolbar.ListImages.Clear
'        '.iml_Toolbar.ListImages.add 1, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Back.gif")
 '       .iml_Toolbar.ListImages.add 2, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Forward.gif")
  '      .iml_Toolbar.ListImages.add 3, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Home.gif")
  '      .iml_Toolbar.ListImages.add 4, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Refresh.gif")
  '      .iml_Toolbar.ListImages.add 5, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Open.gif")
   '     .iml_Toolbar.ListImages.add 6, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Document.gif")
   '     .iml_Toolbar.ListImages.add 7, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Search.gif")
    ''    .iml_Toolbar.ListImages.add 8, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Help.gif")
       ' .iml_Toolbar.ListImages.add 9, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Stop.gif")
      '  .ctrl_Toolbar.UnloadButtons
'        .ctrl_Toolbar.BackColor = &H0&
 '       .ctrl_Toolbar.IconLeft = 90
  '      .ctrl_Toolbar.IconTop = 90
   '     Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
    '    Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
     '   Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        'Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
'        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
 '       Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
 '       Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
  '      Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
'        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
 '       Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
  '      Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
   '     Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
    '    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
     '   Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        '.ctrl_ListObject.DrawMenu
        
 '       .ctrl_Toolbar.SkinPath = App.Path & "\Skins\TreasureChest"
  '      .ctrl_Toolbar.DrawToolbar
  ''      .ctrl_Toolbar.Refresh
   '
    '    .ctrl_Panel.SkinPath = App.Path & "\Skins\TreasureChest"
    '    .ctrl_Panel.DrawPanel
     '
'        .ctrl_PullDownMenu.BackColor = &H0&
 '       .ctrl_PullDownMenu.ForeColor = &HFFFFFF
  '      .ctrl_PullDownMenu.Refresh
   '
    '    .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\TreasureChest"
     '   .ctrl_ChannelBar.SubItemTop = 395
      '  .ctrl_ChannelBar.MouseMoveColor = &H0&
       ' .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        '.ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
       ' .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
'        .ctrl_ChannelBar.DrawMenu
 '
  '      .Line1.BorderColor = &HFFFFFF
   '     .lbl_Statusbar.ForeColor = &HFFFFFF
    '
     '   .pic_Viewport.BackColor = &H304B95
      '  .tbx_Text.BackColor = &H304B95
       ' .tbx_Text.ForeColor = &HFFFFFF
    End With
End Sub

Public Sub ChangeSkinToALPI()
    With frmmain
''        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\ALPI"
'        .ctrl_SkinableForm.BackColor = &H2E2E32
 '       .ctrl_SkinableForm.CaptionTop = 135
  '      .ctrl_SkinableForm.CaptionColor = &H0&
   '     Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
    '
     '   .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\ALPI"
      '  .ctrl_btn_Previous.ForeColor = &HFFFFFF
       ' .ctrl_btn_Previous.LoadSkin
        '.ctrl_btn_Previous.Refresh
'        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\ALPI"
 '       .ctrl_btn_Next.ForeColor = &HFFFFFF
  '      .ctrl_btn_Next.LoadSkin
   '     .ctrl_btn_Next.Refresh
    '    .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\ALPI"
     '   .ctrl_btn_Exit.ForeColor = &HFFFFFF
      '  .ctrl_btn_Exit.LoadSkin
       ' .ctrl_btn_Exit.Refresh
        '
        '.ctrl_ListObject.SkinPath = App.Path & "\Skins\ALPI"
'        .ctrl_ListObject.ForeColor = &HFFFFFF
 '       .ctrl_ListObject.MouseMoveColor = &H0&
  '      .ctrl_ListObject.MouseDownColor = &H0&
   '     .iml_Toolbar.ListImages.Clear
    '    .iml_Toolbar.ListImages.add 1, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Back.gif")
     '   .iml_Toolbar.ListImages.add 2, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Forward.gif")
      ''  .iml_Toolbar.ListImages.add 3, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Home.gif")
        '.iml_Toolbar.ListImages.add 4, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Refresh.gif")
        '.iml_Toolbar.ListImages.add 5, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Open.gif")
'        .iml_Toolbar.ListImages.add 6, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Document.gif")
 '       .iml_Toolbar.ListImages.add 7, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Search.gif")
  '      .iml_Toolbar.ListImages.add 8, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Help.gif")
   '     .iml_Toolbar.ListImages.add 9, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Stop.gif")
    '    .ctrl_Toolbar.UnloadButtons
     '   .ctrl_Toolbar.IconLeft = 90
      '  .ctrl_Toolbar.IconTop = 90
       ' Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        'Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
'        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
 '       Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
  '      Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
   '     Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
    '    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
     '   Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
      '  Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
       ' Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        'Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
'        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
 '       Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
  '      Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
   '     Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
    '    Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
     '   Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
      '  Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
       ' .ctrl_ListObject.DrawMenu
        '
'        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\ALPI"
 '       .ctrl_Toolbar.BackColor = &H2E2E32
  '      .ctrl_Toolbar.DrawToolbar
   '     .ctrl_Toolbar.Refresh
    '
     '   .ctrl_Panel.SkinPath = App.Path & "\Skins\ALPI"
        '.ctrl_Panel.DrawPanel
      '
       ' .ctrl_PullDownMenu.BackColor = &H2E2E32
        '.ctrl_PullDownMenu.ForeColor = &HFFFFFF
        '.ctrl_PullDownMenu.Refresh
        
'        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\ALPI"
 '       .ctrl_ChannelBar.SubItemTop = 395
  '      .ctrl_ChannelBar.MouseMoveColor = &H0&
   '     .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
    '    .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
     '   .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
      '  .ctrl_ChannelBar.DrawMenu
        
       ' .Line1.BorderColor = &HFFFFFF
        '.lbl_Statusbar.ForeColor = &HFFFFFF
        
'        .pic_Viewport.BackColor = &H0&
 '       .tbx_Text.BackColor = &H0&
  '      .tbx_Text.ForeColor = &HFFFFFF
    End With
End Sub

