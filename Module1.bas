Attribute VB_Name = "Module1"
'*************************************************************************************************************************************************************************************************************************
   'Functions and Variables for Fade and Semi Transpararent RTNs
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'*************************************************************************************************************************************************************************************************************************
'*************************************************************************************************************************************************************************************************************************
     'Sleep Declaration and Sample Routine
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Public Sub rtnSleep()
'Sleep for 30 minutes
'Sleep 1,800,000
'Sleep for 1 minute
'Sleep 60,000
'Sleep for 1 second
'Sleep 1,000
'End Sub
'*************************************************************************************************************************************************************************************************************************
Public intMaxPicSize As Integer 'Max Height and Width
Public stFilename As String
Public booReDraw As Boolean 'Redraw sprites
Public pixMainPicture As Picture
Public pixMainPictureWithSprite As Picture 'Used for fadeouts only
Public pixMainPictureWithOutSprite As Picture 'Used for fadeouts only
Public intPlaceX As Integer, intPlaceY As Integer 'Sprite placement
Public intScanX As Integer, intScanY As Integer 'Mask scan/shift values

Public Sub rtnFadeIn()
Dim X As Integer, Y As Integer
Dim pixMaskWhiteDots As Picture, pixMaskBlackDots As Picture, pixMini20x20Mask As Picture
Dim pixOr As Picture, pixAnd As Picture
'This routine does the main fade processing
'It enters with a Main Picture, two Sprite Pictures, and a 20 x 20 pixel mask.
'The two sprite pics have their border colors black and white respectively.
'This is for using the ("paint" and "and") rasterops producing transparency.
'If transparency is selected then this is what transpires.
'The sprites are painted in the usual ("paint" and "and") but before they are painted
'to the main screen this way, they themselves are first painted ("paint" and "and")
'with the little 20 x 20 pixel mask after spreading this mask evenly over the sprites.
'This way what goes unto the main screen will be only those areas of the sprite
'that had a white part of the mask over them. This process is repeated shifting
'the mask each time according to the user selected x,y scan. This way the white
'portions of the mask can be shifted over the entire image producing a very
'pleasing effect. Different masks can be created and different scan values
'can be entered to produce very interesting effects!


With frmMain
Set pixMini20x20Mask = .picMask 'Mask with white producing the transparency
.picMain.Picture = pixMainPicture
Set pixMainPictureWithOutSprite = .picMain.Picture 'used for fadeouts only
.picEdit.AutoSize = False
'Make edit box 20 pixels wider and higher than main screen
'to accomodate the maximum user x,y mask shift
.picEdit.Height = .picMain.Height + 20
.picEdit.Width = .picMain.Width + 20

'Load the 20 x 20 pixel  mask. MUST BE Black and White!!!
'Paint it evenly over the now oversized edit box
For Y = 0 To .picEdit.Height + 20 Step 20
For X = 0 To .picEdit.Width + 20 Step 20
.picEdit.PaintPicture pixMini20x20Mask, X, Y
Next X
Next Y

'Now store the oversized mask, invert and store again
.picEdit.Picture = .picEdit.Image
Set pixMaskWhiteDots = .picEdit.Picture
.picEdit.PaintPicture pixMaskWhiteDots, 0, 0, , , , , , , vbDstInvert
.picEdit.Picture = .picEdit.Image
Set pixMaskBlackDots = .picEdit.Picture

'Main Loop
.picEdit.AutoSize = True
For Y = -intScanY + 1 To 0 'Mask user Y shift
For X = -intScanX + 1 To 0 'Mask user X shift
                             
                             '1st "OR" sprite
    'Get pix to fade
    'Paint the mask with white dots and black bkg
        'If solid selected
        If .Option1(1).Value = True Then
       .picEdit.Picture = .picToFadeIn
        'If transparency selected
        Else
       .picEdit.Picture = .picSpriteBlkBkg.Picture
        End If
       .picEdit.PaintPicture pixMaskWhiteDots, X, Y, , , , , , , vbSrcAnd
       Set pixOr = .picEdit.Image
     'This image is now:
    'Background = Black
    'Dots placed according to mask dots position
    'Dots are the same colors as original image
                              
                              '2nd AND" sprite
    'Get pix to fade
    'Paint the mask with black dots and white bkg
        'If solid selected
        If .Option1(1).Value = True Then
       .picEdit.Picture = .picToFadeIn
        'If transparency selected
        Else
       .picEdit.Picture = .picSpriteWhtBkg.Picture
        End If
       .picEdit.PaintPicture pixMaskBlackDots, X, Y, , , , , , , vbSrcPaint 'paint mask on 1/2 sprite
       Set pixAnd = .picEdit.Image 'save 1/2 sprite
    'This image is now:
    'Background = copy of original image
    'Dots placed according to mask dots position
    'Dots are White
                              
                              'Main paint sprites 1/2 & 2/2
    .picMain.PaintPicture pixOr, Val(intPlaceX), Val(intPlaceY), , , 0, 0, , , vbSrcPaint
    .picMain.PaintPicture pixAnd, Val(intPlaceX), Val(intPlaceY), , , 0, 0, , , vbSrcAnd
    .picMain.Refresh
     DoEvents
Sleep .scrllSleep.Value
Next X
Next Y
.picMain.Picture = .picMain.Image
Set pixMainPictureWithSprite = .picMain.Picture
.picEdit.AutoSize = False

Set pixOr = LoadPicture("") 'Clear
Set pixAnd = LoadPicture("") 'Clear
          .picEdit = LoadPicture("")
End With
End Sub

Public Sub rtnFadeOut()
Dim X As Integer, Y As Integer
Dim pixMaskWhiteDots As Picture, pixMaskBlackDots As Picture, pixMini20x20Mask As Picture
Dim pixOr As Picture, pixAnd As Picture
'This routine is for fading out only
'It is basically the same as the fade in rtn with the appropiate pictures swapped

With frmMain
Set pixMini20x20Mask = .picMask 'Mask with white producing the transparency
.picMain.Picture = pixMainPictureWithSprite
.picEdit.AutoSize = False
'Make edit box 20 pixels wider and higher than main screen
'to accomodate the maximum user x,y mask shift
.picEdit.Height = .picMain.Height + 20
.picEdit.Width = .picMain.Width + 20

'Load the 20 x 20 pixel  mask. MUST BE Black and White!!!
'Paint it evenly over the now oversized edit box
For Y = 0 To .picEdit.Height + 20 Step 20
For X = 0 To .picEdit.Width + 20 Step 20
.picEdit.PaintPicture pixMini20x20Mask, X, Y
Next X
Next Y

'Now store the oversized mask, invert and store again
.picEdit.Picture = .picEdit.Image
Set pixMaskWhiteDots = .picEdit.Picture
.picEdit.PaintPicture pixMaskWhiteDots, 0, 0, , , , , , , vbDstInvert
.picEdit.Picture = .picEdit.Image
Set pixMaskBlackDots = .picEdit.Picture


.picEdit.AutoSize = True
For Y = -intScanY + 1 To 0 'Mask user Y shift
For X = -intScanX + 1 To 0 'Mask user X shift
                             '1st picture
    'Get pix to dissolve
    'Paint the mask with white dots and black bkg
       .picEdit.Picture = pixMainPictureWithOutSprite
       .picEdit.PaintPicture pixMaskWhiteDots, X, Y, , , , , , , vbSrcAnd
       Set pixOr = .picEdit.Image
     'This image is now:
    'Background = Black
    'Dots placed according to mask dots position
    'Dots are the same colors as original image
                              
                              '2nd picture
    'Get pix to dissolve
    'Paint the mask with black dots and white bkg
       .picEdit.Picture = pixMainPictureWithOutSprite
       .picEdit.PaintPicture pixMaskBlackDots, X, Y, , , , , , , vbSrcPaint 'paint mask on 1/2 sprite
       Set pixAnd = .picEdit.Image 'save 1/2 sprite
    'This image is now:
    'Background = copy of original image
    'Dots placed according to mask dots position
    'Dots are White
                              
                              'Main paint sprites 1/2 & 2/2
    .picMain.PaintPicture pixOr, Val(intPlaceX), Val(intPlaceY), , , 0, 0, , , vbSrcPaint
    .picMain.PaintPicture pixAnd, Val(intPlaceX), Val(intPlaceY), , , 0, 0, , , vbSrcAnd
    .picMain.Refresh
     DoEvents
Sleep .scrllSleep.Value
Next X
Next Y
.picMain.Picture = .picMain.Image
Set pixMainPictureWithSprite = .picMain.Picture
.picEdit.AutoSize = False

Set pixOr = LoadPicture("") 'Clear
Set pixAnd = LoadPicture("") 'Clear
          .picEdit = LoadPicture("")
End With
End Sub


Public Sub rtnShowPictureDraw()
Dim X, Y, xx, Pointer, Pointer2 As Integer
Dim rTrace, gTrace, bTrace As Byte
Dim rDraw, gDraw, bDraw As Byte
Dim w As Long
Dim colortrace As Long
Dim arrMain(4000, 1) As Integer   'inArrUnfinishedPixels(cell,sides,xy)
Dim arrNewAdjPxls(4000, 1) As Integer   'inArrUnfinishedPixels(cell,sides,xy)
Dim arrAdjPxls(4000, 1) As Integer   'inArrUnfinishedPixels(cell,sides,xy)
Dim color As Long
Const cX = 0
Const cY = 1

'This routine paints our sprite backgrounds for transparency.
'It logs the color of the topleft corner pixel
'of our sprite as "colortrace"
'Then it checks each perimeter pixel against this "trace" color.
'Of those matching, their adjacent pixels are checked,
'and of those matching their adjacents are checked and do forth.
'All matching pixels are painted on two seperate picture boxes, black
'on one pix and white on the other thus producing our transpareny sprite background.


With frmMain
.picHidden1.Picture = .picToFadeIn.Picture 'Autosize picture according to pix to fade in
.picSpriteBlkBkg.Picture = .picToFadeIn.Picture '"OR" sprite
.picSpriteWhtBkg.Picture = .picToFadeIn.Picture '"AND" sprite

'..............................................................
'..............................................................
'..............................................................
'.................MAIN DRAWING ROUTINE.........................
'..............................................................
'..............................................................
'..............................................................
'Make or topleft corner pixel our trace color
color = GetPixel(.picHidden1.hDC, 0, 0)
'Purple is our "processed" marker color
If color = RGB(255, 0, 255) Then Exit Sub
colortrace = color
'Check each perimeter pixel
Y = 0
For X = 0 To .picHidden1.ScaleWidth - 1
GoSub lblStart
Next X
Y = .picHidden1.ScaleHeight - 1
For X = 0 To .picHidden1.ScaleWidth - 1
GoSub lblStart
Next X
X = 0
For Y = 0 To .picHidden1.ScaleHeight - 1
GoSub lblStart
Next Y
X = .picHidden1.ScaleWidth - 1
For Y = 0 To .picHidden1.ScaleHeight - 1
GoSub lblStart
Next Y
.picSpriteBlkBkg.Picture = .picSpriteBlkBkg.Image
.picSpriteWhtBkg.Picture = .picSpriteWhtBkg.Image
'Finished
Exit Sub

lblStart:
color = GetPixel(.picHidden1.hDC, X, Y)
If color = RGB(255, 0, 255) Then GoTo lblReturn 'pixel already processed, check next pixel
If color <> colortrace Then GoTo lblReturn 'Didn't match our corner "trace" pixel, check next pixel
    
'We have a pixel that matched our colortrace
SetPixel .picHidden1, X, Y, RGB(255, 0, 255) 'Mark the pixel
SetPixel .picSpriteBlkBkg, X, Y, RGB(0, 0, 0) 'Paint our "OR" sprite respectively
SetPixel .picSpriteWhtBkg, X, Y, RGB(255, 255, 255) 'Paint our "AND" sprite respectively
arrNewAdjPxls(1, cX) = X: arrNewAdjPxls(1, cY) = Y ' Save pixel's xy temporarily to array
Pointer = 1 'How many pixels are in array that still need their adjacents checked

    
'..............................................................
'.................MAIN LOOP....................................
'..............................................................
Do While Pointer > 0 'Do while there are still adjacents pixels left to check
Pointer2 = Pointer: Pointer = 0 'Reset
        
   'Swap arrays
    For xx = 1 To Pointer2
    arrAdjPxls(xx, cX) = arrNewAdjPxls(xx, cX)
    arrAdjPxls(xx, cY) = arrNewAdjPxls(xx, cY)
    Next xx
  
   'Check the adjacent pixels of the pixels
   'that were processed in the last pass.
    For xx = 1 To Pointer2
   'Get xy of previously processed pixel from array
    X = arrAdjPxls(xx, cX): Y = arrAdjPxls(xx, cY)
   'Check adjacent Right pixel
    color = GetPixel(.picHidden1.hDC, X + 1, Y)
        If color = colortrace Then
        SetPixel .picHidden1.hDC, X + 1, Y, RGB(255, 0, 255)
       'SetPixel .picMain.hDC, X + 1, Y, RGB(rDraw, gDraw, bDraw)
        SetPixel .picSpriteBlkBkg.hDC, X + 1, Y, RGB(0, 0, 0) 'Set First Pixel on main window
        SetPixel .picSpriteWhtBkg.hDC, X + 1, Y, RGB(255, 255, 255) 'Set First Pixel on main window
        Pointer = Pointer + 1  'Set pointer
        arrNewAdjPxls(Pointer, cX) = X + 1:    arrNewAdjPxls(Pointer, cY) = Y
        End If
   'Check adjacent Bottom pixel
    color = GetPixel(.picHidden1.hDC, X, Y + 1)
        If color = colortrace Then
        SetPixel .picHidden1.hDC, X, Y + 1, RGB(255, 0, 255)
       'SetPixel .picMain.hDC, X, Y + 1, RGB(rDraw, gDraw, bDraw)
        SetPixel .picSpriteBlkBkg.hDC, X, Y + 1, RGB(0, 0, 0) 'Set First Pixel on main window
        SetPixel .picSpriteWhtBkg.hDC, X, Y + 1, RGB(255, 255, 255) 'Set First Pixel on main window
        Pointer = Pointer + 1  'Set pointer
        arrNewAdjPxls(Pointer, cX) = X:     arrNewAdjPxls(Pointer, cY) = Y + 1
        End If
   'Check adjacent BottomRight pixel
    color = GetPixel(.picHidden1.hDC, X + 1, Y + 1)
        If color = colortrace Then
        SetPixel .picHidden1.hDC, X + 1, Y + 1, RGB(255, 0, 255)
       'SetPixel .picMain.hDC, X + 1, Y + 1, RGB(rDraw, gDraw, bDraw)
        SetPixel .picSpriteBlkBkg.hDC, X + 1, Y + 1, RGB(0, 0, 0) 'Set First Pixel on main window
        SetPixel .picSpriteWhtBkg.hDC, X + 1, Y + 1, RGB(255, 255, 255) 'Set First Pixel on main window
        Pointer = Pointer + 1  'Set pointer
        arrNewAdjPxls(Pointer, cX) = X + 1:    arrNewAdjPxls(Pointer, cY) = Y + 1
        End If
   'Check adjacent TopLeft pixel
    color = GetPixel(.picHidden1.hDC, X - 1, Y - 1)
            If color = colortrace Then
            SetPixel .picHidden1.hDC, X - 1, Y - 1, RGB(255, 0, 255)
           'SetPixel .picMain.hDC, X - 1, Y - 1, RGB(rDraw, gDraw, bDraw)
            SetPixel .picSpriteBlkBkg.hDC, X - 1, Y - 1, RGB(0, 0, 0) 'Set First Pixel on main window
            SetPixel .picSpriteWhtBkg.hDC, X - 1, Y - 1, RGB(255, 255, 255) 'Set First Pixel on main window
            Pointer = Pointer + 1  'Set pointer
            arrNewAdjPxls(Pointer, cX) = X - 1:   arrNewAdjPxls(Pointer, cY) = Y - 1
            End If
   'Check adjacent Left pixel
    color = GetPixel(.picHidden1.hDC, X - 1, Y)
        If color = colortrace Then
        SetPixel .picHidden1.hDC, X - 1, Y, RGB(255, 0, 255)
       'SetPixel .picMain.hDC, X - 1, Y, RGB(rDraw, gDraw, bDraw)
        SetPixel .picSpriteBlkBkg.hDC, X - 1, Y, RGB(0, 0, 0)  'Set First Pixel on main window
        SetPixel .picSpriteWhtBkg.hDC, X - 1, Y, RGB(255, 255, 255)  'Set First Pixel on main window
        Pointer = Pointer + 1  'Set pointer
        arrNewAdjPxls(Pointer, cX) = X - 1:    arrNewAdjPxls(Pointer, cY) = Y
        End If
   'Check adjacent Top pixel
    color = GetPixel(.picHidden1.hDC, X, Y - 1)
        If color = colortrace Then
        SetPixel .picHidden1.hDC, X, Y - 1, RGB(255, 0, 255)
       'SetPixel .picMain.hDC, X, Y - 1, RGB(rDraw, gDraw, bDraw)
        SetPixel .picSpriteBlkBkg.hDC, X, Y - 1, RGB(0, 0, 0) 'Set First Pixel on main window
        SetPixel .picSpriteWhtBkg.hDC, X, Y - 1, RGB(255, 255, 255)  'Set First Pixel on main window
        Pointer = Pointer + 1  'Set pointer
        arrNewAdjPxls(Pointer, cX) = X:     arrNewAdjPxls(Pointer, cY) = Y - 1
        End If
   'Check adjacent BottomLeft pixel
    color = GetPixel(.picHidden1.hDC, X - 1, Y + 1)
        If color = colortrace Then
        SetPixel .picHidden1.hDC, X - 1, Y + 1, RGB(255, 0, 255)
       'SetPixel .picMain.hDC, X - 1, Y + 1, RGB(rDraw, gDraw, bDraw)
        SetPixel .picSpriteBlkBkg.hDC, X - 1, Y + 1, RGB(0, 0, 0) 'Set First Pixel on main window
        SetPixel .picSpriteWhtBkg.hDC, X - 1, Y + 1, RGB(255, 255, 255) 'Set First Pixel on main window
        Pointer = Pointer + 1  'Set pointer
        arrNewAdjPxls(Pointer, cX) = X - 1:    arrNewAdjPxls(Pointer, cY) = Y + 1
        End If
   'Check adjacent TopRight pixel
    color = GetPixel(.picHidden1.hDC, X + 1, Y - 1)
        If color = colortrace Then
        SetPixel .picHidden1.hDC, X + 1, Y - 1, RGB(255, 0, 255)
       'SetPixel .picMain.hDC, X + 1, Y - 1, RGB(rDraw, gDraw, bDraw)
        SetPixel .picSpriteBlkBkg.hDC, X + 1, Y - 1, RGB(0, 0, 0) 'Set First Pixel on main window
        SetPixel .picSpriteWhtBkg.hDC, X + 1, Y - 1, RGB(255, 255, 255) 'Set First Pixel on main window
        Pointer = Pointer + 1  'Set pointer
        arrNewAdjPxls(Pointer, cX) = X + 1:     arrNewAdjPxls(Pointer, cY) = Y - 1
        End If
    Next xx
   
   .picSpriteBlkBkg.Refresh
   .picSpriteWhtBkg.Refresh
    Loop
    '..............................................................
    '.................END MAIN LOOP.................................
    '..............................................................

lblReturn:
Return
'..............................................................
'..............................................................
'..............................................................
'.................END MAIN DRAWING ROUTINE.........................
'..............................................................
'..............................................................
'..............................................................

End With
End Sub

