
Public Class line

   Public X1 As Double
   Public Y1 As Double
   Public Z1 As Double
   Public X2 As Double
   Public Y2 As Double
   Public Z2 As Double

   Public Sub New(ByVal pdblX1 As Double, ByVal pdblY1 As Double, ByVal pdblZ1 As Double, ByVal pdblX2 As Double, ByVal pdblY2 As Double, ByVal pdblZ2 As Double)

      X1 = pdblX1
      Y1 = pdblY1
      Z1 = pdblZ1
      X2 = pdblX2
      Y2 = pdblY2
      Z2 = pdblZ2

   End Sub

End Class

Public Class frmMain

   Private mobjFormBitmap As Bitmap
   Private mobjBitmapGraphics As Graphics
   Private mintFormWidth As Integer
   Private mintFormHeight As Integer
   Private mobjLines As List(Of line)
   Private mdblXDisplacement As Double = 0
   Private mdblYDisplacement As Double = 0
   Private mdblZDisplacement As Double = 0
   Private mdblXRotation As Double = 0
   Private mdblYRotation As Double = 0
   Private mdblZRotation As Double = 0

   Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

      Static blnDoneOnce As Boolean = False

      If Not blnDoneOnce Then
         blnDoneOnce = True
         mintFormWidth = Me.Width
         mintFormHeight = Me.Height
         mobjFormBitmap = New Bitmap(mintFormWidth, mintFormHeight, Me.CreateGraphics())
         mobjBitmapGraphics = Graphics.FromImage(mobjFormBitmap)
         mobjBitmapGraphics.FillRectangle(Brushes.White, 0, 0, mintFormWidth, mintFormHeight)
         mobjLines = New List(Of line)
         LoadLines()
         DrawView()
      End If

   End Sub

   Private Sub frmMain_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

      e.Graphics.DrawImage(mobjFormBitmap, 0, 0)

   End Sub


   Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

      Select Case e.KeyCode
         Case Keys.Up
            mdblZDisplacement = 10
         Case Keys.Down
            mdblZDisplacement = -10
         Case Keys.Right
            mdblYRotation = Math.PI / 64
         Case Keys.Left
            mdblYRotation = -Math.PI / 64
         Case Keys.W
            mdblYDisplacement = 10
         Case Keys.Z
            mdblYDisplacement = -10
         Case Keys.A
            mdblXDisplacement = -10
         Case Keys.S
            mdblXDisplacement = 10
         Case Keys.I
            mdblXRotation = -Math.PI / 64
         Case Keys.M
            mdblXRotation = Math.PI / 64
         Case Keys.K
            mdblZRotation = -Math.PI / 64
         Case Keys.J
            mdblZRotation = Math.PI / 64
      End Select

      DrawView()

      mdblZDisplacement = 0
      mdblZDisplacement = 0
      mdblYRotation = 0
      mdblYRotation = 0
      mdblYDisplacement = 0
      mdblYDisplacement = 0
      mdblXDisplacement = 0
      mdblXDisplacement = 0
      mdblXRotation = 0
      mdblXRotation = 0
      mdblZRotation = 0
      mdblZRotation = 0

   End Sub

   Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)

      ' to remove the flickering

   End Sub

   Private Sub LoadLines()

      For intXindex As Int16 = 0 To 0
         For intYindex As Int16 = 0 To 0
            For intZindex As Int16 = 0 To 0
               mobjLines.Add(New line((intXindex * 90), (intYindex * 90), 1000 + (intZindex * 90), (intXindex * 90) + 60, (intYindex * 90), 1000 + (intZindex * 90)))
               mobjLines.Add(New line((intXindex * 90) + 60, (intYindex * 90), 1000 + (intZindex * 90), (intXindex * 90) + 60, (intYindex * 90) + 60, 1000 + (intZindex * 90)))
               mobjLines.Add(New line((intXindex * 90) + 60, (intYindex * 90) + 60, 1000 + (intZindex * 90), (intXindex * 90), (intYindex * 90) + 60, 1000 + (intZindex * 90)))
               mobjLines.Add(New line((intXindex * 90), (intYindex * 90) + 60, 1000 + (intZindex * 90), (intXindex * 90), (intYindex * 90), 1000 + (intZindex * 90)))
               mobjLines.Add(New line((intXindex * 90), (intYindex * 90), 1000 + (intZindex * 90) + 60, (intXindex * 90) + 60, (intYindex * 90), 1000 + (intZindex * 90) + 60))
               mobjLines.Add(New line((intXindex * 90) + 60, (intYindex * 90), 1000 + (intZindex * 90) + 60, (intXindex * 90) + 60, (intYindex * 90) + 60, 1000 + (intZindex * 90) + 60))
               mobjLines.Add(New line((intXindex * 90) + 60, (intYindex * 90) + 60, 1000 + (intZindex * 90) + 60, (intXindex * 90), (intYindex * 90) + 60, 1000 + (intZindex * 90) + 60))
               mobjLines.Add(New line((intXindex * 90), (intYindex * 90) + 60, 1000 + (intZindex * 90) + 60, (intXindex * 90), (intYindex * 90), 1000 + (intZindex * 90) + 60))
               mobjLines.Add(New line((intXindex * 90), (intYindex * 90), 1000 + (intZindex * 90), (intXindex * 90), (intYindex * 90), 1000 + (intZindex * 90) + 60))
               mobjLines.Add(New line((intXindex * 90) + 60, (intYindex * 90), 1000 + (intZindex * 90), (intXindex * 90) + 60, (intYindex * 90), 1000 + (intZindex * 90) + 60))
               mobjLines.Add(New line((intXindex * 90) + 60, (intYindex * 90) + 60, 1000 + (intZindex * 90), (intXindex * 90) + 60, (intYindex * 90) + 60, 1000 + (intZindex * 90) + 60))
               mobjLines.Add(New line((intXindex * 90), (intYindex * 90) + 60, 1000 + (intZindex * 90), (intXindex * 90), (intYindex * 90) + 60, 1000 + (intZindex * 90) + 60))
            Next
         Next
      Next

   End Sub

   Private Sub DrawView()

      Dim objLine As line

      mobjBitmapGraphics.FillRectangle(Brushes.White, 0, 0, mintFormWidth, mintFormHeight)

      For intLineCount As Int16 = 0 To mobjLines.Count - 1
         objLine = mobjLines(intLineCount)
         If TransformLine(objLine) Then
            DrawLine(objLine)
         End If
         mobjLines(intLineCount) = objLine
      Next

      Me.Invalidate()
      Application.DoEvents()

   End Sub

   Private Function TransformLine(ByRef pobjLine As line) As Boolean

      Dim dblX1 As Double
      Dim dblY1 As Double
      Dim dblZ1 As Double
      Dim dblX2 As Double
      Dim dblY2 As Double
      Dim dblZ2 As Double
      Dim blnPoint1Ok As Boolean
      Dim blnPoint2Ok As Boolean

      dblX1 = pobjLine.X1
      dblY1 = pobjLine.Y1
      dblZ1 = pobjLine.Z1

      dblX2 = pobjLine.X2
      dblY2 = pobjLine.Y2
      dblZ2 = pobjLine.Z2

      blnPoint1Ok = TransformPoint(dblX1, dblY1, dblZ1)
      blnPoint2Ok = TransformPoint(dblX2, dblY2, dblZ2)

      pobjLine.X1 = dblX1
      pobjLine.Y1 = dblY1
      pobjLine.Z1 = dblZ1

      pobjLine.X2 = dblX2
      pobjLine.Y2 = dblY2
      pobjLine.Z2 = dblZ2

      Return blnPoint1Ok And blnPoint2Ok

   End Function

   Private Function TransformPoint(ByRef pdblX As Double, ByRef pdblY As Double, ByRef pdblZ As Double)

      Dim dblX As Double
      Dim dblY As Double
      Dim dblZ As Double

      ' rotation around x
      dblX = pdblX
      dblY = pdblY * Math.Cos(mdblXRotation) + pdblZ * Math.Sin(mdblXRotation)
      dblZ = -pdblY * Math.Sin(mdblXRotation) + pdblZ * Math.Cos(mdblXRotation)

      pdblX = dblX
      pdblY = dblY
      pdblZ = dblZ

      ' rotation around y
      dblX = pdblX * Math.Cos(mdblYRotation) - pdblZ * Math.Sin(mdblYRotation)
      dblY = pdblY
      dblZ = pdblX * Math.Sin(mdblYRotation) + pdblZ * Math.Cos(mdblYRotation)

      pdblX = dblX
      pdblY = dblY
      pdblZ = dblZ

      ' rotation around z
      dblX = pdblX * Math.Cos(mdblZRotation) + pdblY * Math.Sin(mdblZRotation)
      dblY = -pdblX * Math.Sin(mdblZRotation) + pdblY * Math.Cos(mdblZRotation)
      dblZ = pdblZ

      pdblX = dblX
      pdblY = dblY
      pdblZ = dblZ

      pdblX = pdblX - mdblXDisplacement
      pdblY = pdblY - mdblYDisplacement
      pdblZ = pdblZ - mdblZDisplacement

      If pdblZ > 0 Then
         Return True
      Else
         Return False
      End If

   End Function

   Private Sub DrawLine(ByVal pobjLine As line)

      Dim sglX1Projected As Single
      Dim sglY1Projected As Single
      Dim sglX2Projected As Single
      Dim sglY2Projected As Single

      sglX1Projected = (mintFormWidth / 2) + (pobjLine.X1 * 700 / pobjLine.Z1)
      sglY1Projected = (mintFormHeight / 2) - (pobjLine.Y1 * 700 / pobjLine.Z1)

      sglX2Projected = (mintFormWidth / 2) + (pobjLine.X2 * 700 / pobjLine.Z2)
      sglY2Projected = (mintFormHeight / 2) - (pobjLine.Y2 * 700 / pobjLine.Z2)

      mobjBitmapGraphics.DrawLine(Pens.Black, sglX1Projected, sglY1Projected, sglX2Projected, sglY2Projected)

   End Sub


End Class



