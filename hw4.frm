VERSION 5.00
Begin VB.Form form1 
   Caption         =   "numberofgene"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   12555
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   6720
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   12135
   End
   Begin VB.CommandButton read 
      Caption         =   "讀檔"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton summit 
      Caption         =   "確定"
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox kvalue 
      Height          =   495
      Left            =   8760
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton score 
      Caption         =   "score"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton tvalue 
      Caption         =   "tvalue"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox numberofgene 
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox file 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "colon.txt"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "K Value"
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Gene"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "檔名"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filename As String
Dim g As Double
Dim k As Double
Dim Dblgene(2000, 62) As Double
Dim InverseDblgene(62, 2000) As Double
Dim genenumber(1999) As Double
Dim genescore(1999) As Double
Dim tvaluearray(1, 1999) As Double

Private Sub Text1_Change()

End Sub

Private Sub file_Change()
filename = file.Text
End Sub

Private Sub kvalue_Change()
k = CDbl(kvalue.Text)
End Sub

Private Sub numberofgene_Change()
g = CDbl(numberofgene.Text)
End Sub

Private Sub read_Click()
List1.Clear
Dim counter As Double
Dim class() As String
Dim Egene() As String
Dim ten As Double
counter = -1
ten = 10

Open App.Path & "\" + filename For Input As #1
Do While Not EOF(1)

Line Input #1, tmpline
    If counter = 0 Then
        class = Split(tmpline, ",")
        For i = 1 To 62
        Dblgene(0, i) = CDbl(class(i))
        Next i

    Else

        If counter <> -1 Then
            Dim gene() As String
            gene = Split(tmpline, ",")
            
            For i = 0 To 62
                gene(i) = Trim(gene(i))
                If InStr(gene(i), "E") <> 0 Then
                    Egene = Split(gene(i), "E")
                    Dblgene(counter, i) = CDbl(CDbl(Egene(0)) * (ten ^ CDbl(-Right(Egene(1), 1))))
                Else
                    Dblgene(counter, i) = CDbl(gene(i))
                End If
            Next i
            
            
        End If
    End If

    counter = counter + 1



Loop
Close #1

'For i = 0 To 2000
'For j = 0 To 62
'List1.AddItem Dblgene(i, j)
'Next j
'Next i
List1.AddItem "Read success"
'List1.AddItem Dblgene(1, 3)
'List1.AddItem Dblgene(2, 3)
End Sub

Private Sub score_Click()
List1.Clear
List1.AddItem "-----Score-----"
'測試sortrndi
'Dim test1(2) As Double
'Dim test2(2) As Double
'Dim out As String
'test2(0) = 1
'test2(1) = 2
'test2(2) = 3
'test1(0) = 0.5
'test1(1) = 0.7
'test1(2) = 0.3
'For i = 0 To 2
'List1.AddItem CStr(test2(i))
'List1.AddItem CStr(test1(i))
'Next i
'out = sortrndi(test2, test1)
'For i = 0 To 2
'List1.AddItem CStr(test2(i))
'List1.AddItem CStr(test1(i))
'Next i


Dim scorenum As Double
Dim one As Double
Dim two As Double
Dim sortout As String


For i = 1 To 2000
scorenum = 0
For j = 1 To 62
    If (Dblgene(0, j) = 1) Then
    one = Dblgene(i, j)
    For k = 1 To 62
        If (Dblgene(0, k) = 2) Then
        two = Dblgene(i, k)
            If (one > two) Then
            scorenum = scorenum + 1
            End If
        End If
    Next k
    End If
Next j
genenumber(i - 1) = Dblgene(i, 0)
genescore(i - 1) = scorenum

'List1.AddItem genenumber(i - 1)
'List1.AddItem genescore(i - 1)
Next i

'List1.AddItem "--------------------------------"
sortout = sortrnd(genenumber, genescore)

For i = 0 To 1999
'List1.AddItem genenumber(i)
'List1.AddItem genescore(i)
List1.AddItem "Gene seq  " & CStr(genenumber(i)) & vbTab & "Score =   " & CStr(genescore(i))
Next i

'GoTo scoreend
'scoreend:
End Sub
'大到小
Static Function sortrnd(ByRef tempdataindex() As Double, ByRef temprndarray() As Double)

Dim tmp As Double
Dim tmpindex As Double

For i = 0 To UBound(tempdataindex)
    For j = i To UBound(tempdataindex)
        If temprndarray(i) < temprndarray(j) Then
            tmp = temprndarray(i)
            temprndarray(i) = temprndarray(j)
            temprndarray(j) = tmp
            
            tmpindex = tempdataindex(i)
            tempdataindex(i) = tempdataindex(j)
            tempdataindex(j) = tmpindex
        End If
    Next j
Next i


sortrnd = "sortrnd"
End Function
'小到大
Static Function sortrndi(ByRef tempdataindex() As Double, ByRef temprndarray() As Double)

Dim tmp As Double
Dim tmpindex As Double

For i = 0 To UBound(tempdataindex)
    For j = i To UBound(tempdataindex)
        If temprndarray(i) > temprndarray(j) Then
            tmp = temprndarray(i)
            temprndarray(i) = temprndarray(j)
            temprndarray(j) = tmp
            
            tmpindex = tempdataindex(i)
            tempdataindex(i) = tempdataindex(j)
            tempdataindex(j) = tmpindex
        End If
    Next j
Next i


sortrndi = "sortrnd"
End Function

Static Function leaveoneout(ByVal obs As Double)
Dim tempobs As Double
Dim tempdist(60) As Double
Dim tempgnum(60) As Double
Dim tempgene() As Double
Dim kgnum() As Double
Dim kdist() As Double
Dim tempk As Double
Dim gnumcounter As Double
Dim temppredclass As Double
Dim sortrndout As String
Dim ans As String
ReDim tempgene(g - 1)
ReDim kgnum(k - 1)
ReDim kdist(k - 1)
For i = 0 To UBound(tempgene)
tempgene(i) = genenumber(i)
Next i
gnumcounter = 0
tempobs = obs
tempk = k
ans = ""


For i = 1 To 62
If (i = tempobs) Then
GoTo forend
End If
tempgnum(gnumcounter) = i
gnumcounter = gnumcounter + 1
forend:
Next i

For i = 0 To 60
tempdist(i) = distance(tempobs, tempgnum(i), tempgene)
Next i

sortrndout = sortrndi(tempgnum, tempdist)

For i = 0 To (tempk - 1)
kgnum(i) = tempgnum(i) '這是obs
kdist(i) = tempdist(i)

'List1.AddItem kgnum(i)
'List1.AddItem kdist(i)
Next i

'List1.AddItem "---------------------"

temppredclass = weightvote(kgnum, kdist)

If (temppredclass = Dblgene(0, tempobs)) Then
ans = "y"
Else
ans = "n"
End If



leaveoneout = ans
End Function
'ok
Static Function weightvote(ByRef tempkobs() As Double, ByRef tempkdist() As Double)
Dim predclass As Double
Dim onecounter As Double
Dim twocounter As Double
Dim kobs() As Double
Dim kdist() As Double
kobs = tempkobs()
kdist = tempkdist()
onecounter = 0
twocounter = 0
predclass = 0


For i = 0 To UBound(kobs)
If (Dblgene(0, kobs(i)) = 1) Then
'debug'If (1 = 1) Then
onecounter = onecounter + (1 / (tempkdist(i)) ^ 2)
Else
twocounter = twocounter + (1 / (tempkdist(i)) ^ 2)
End If
Next i


If (onecounter > twocounter) Then
predclass = 1
Else
predclass = 2
End If


weightvote = predclass
End Function

'ok
Private Sub summit_Click()
List1.Clear
'GoTo summitend

Dim g_now As Double
Dim k_now As Double
Dim ycounter As Double
Dim correctrate As Double
Dim tempans As String
g_now = g
k_now = k
ycounter = 0


For i = 1 To 62
tempans = leaveoneout(i)
If (tempans = "y") Then
ycounter = ycounter + 1
End If
Next i
correctrate = (ycounter / 62)
List1.AddItem correctrate

summitend:

'debugweightvote
'Dim test1(1) As Double
'Dim test2(1) As Double
'Dim ans As Double
'test1(0) = 1
'test1(1) = 2
'test2(0) = 2
'test2(1) = 4
'ans = weightvote(test1, test2)

'把Dblgene反轉
'For i = 0 To 62
'For j = 0 To 2000
'InverseDblgene(i, j) = Dblgene(j, i)
'Next j
'Next i

'測試 distance
'Dim gene(1) As Double
'Dim dist As Double
'gene(0) = 5
'gene(1) = 6
'dist = distance(1, 2, gene)
'List1.AddItem dist




End Sub
'ok
Static Function distance(ByVal xindex As Double, ByVal yindex As Double, ByRef attrarray() As Double)
Dim dimnumber As Double
Dim xydistance As Double
Dim totalsum As Double
Dim tempattrarray() As Double
Dim xarray() As Double
Dim yarray() As Double
xydistance = 0
totalsum = 0
tempattrarray() = attrarray()
dimnumber = UBound(tempattrarray) + 1
ReDim xarray(UBound(tempattrarray))
ReDim yarray(UBound(tempattrarray))

For i = 0 To UBound(tempattrarray)
xarray(i) = CDbl(Dblgene(tempattrarray(i), xindex))
yarray(i) = CDbl(Dblgene(tempattrarray(i), yindex))
Next i

'For i = 0 To UBound(tempattrarray)
'List1.AddItem xarray(i)
'Next i
'List1.AddItem ""
'For i = 0 To UBound(tempattrarray)
'List1.AddItem yarray(i)
'Next i


For i = 0 To UBound(tempattrarray)
totalsum = totalsum + ((xarray(i) - yarray(i)) ^ 2)
Next i

xydistance = (totalsum ^ (1 / 2))

distance = xydistance
End Function

Private Sub tvalue_Click()
List1.Clear
List1.AddItem "-----Tvalue-----"
'declare variable
Dim counter As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim class() As String
Dim gene() As String
Dim Dblgene(62) As Double
Dim Egene() As String
Dim sortout As String
Dim ten As Double
Dim n1array(62) As Integer
Dim n2array(62) As Integer
Dim sum1 As Double
Dim sum2 As Double
Dim x1 As Double
Dim x2 As Double
Dim vsum1 As Double
Dim vsum2 As Double
Dim var1 As Double
Dim var2 As Double
Dim t As Double
Dim output, ke, it
Dim outputarray(2000) As Double
Set output = CreateObject("Scripting.Dictionary")

'assign value
counter = -1
n1 = 0
n2 = 0
ten = 10

Open App.Path & "\" + file For Input As #1

Do While Not EOF(1)
    Line Input #1, tmpline
    
    sum1 = 0
    sum2 = 0
    vsum1 = 0
    vsum2 = 0

    If counter = 0 Then
        class = Split(tmpline, ",")
        For i = 1 To 62
            If class(i) = "1" Then
                n1array(i) = i
                n1 = n1 + 1
            Else
                n2array(i) = i
            End If
        Next i
        n2 = 62 - n1
    
    Else
        If counter <> -1 Then
            gene = Split(tmpline, ",")
            For i = 0 To 62
                gene(i) = Trim(gene(i))
                If InStr(gene(i), "E") <> 0 Then
                    Egene = Split(gene(i), "E")
                    Dblgene(i) = CDbl(CDbl(Egene(0)) * (ten ^ CDbl(-Right(Egene(1), 1))))
                Else
                    Dblgene(i) = CDbl(gene(i))
                End If
            Next
            
            For i = 1 To 62
                If n1array(i) <> 0 Then
                    sum1 = sum1 + Dblgene(i)
                End If
                
                If n2array(i) <> 0 Then
                    sum2 = sum2 + Dblgene(i)
                End If
            Next
            
            x1 = (sum1 / CDbl(n1))
            x2 = (sum2 / CDbl(n2))
            
            For i = 1 To 62
                If n1array(i) <> 0 Then
                    vsum1 = vsum1 + (Dblgene(i) - x1) ^ 2
                End If
                
                If n2array(i) <> 0 Then
                    vsum2 = vsum2 + (Dblgene(i) - x2) ^ 2
                End If
            Next
            
            var1 = vsum1 / (n1 - 1)
            var2 = vsum2 / (n2 - 1)
            
            t = (x1 - x2) / ((var1 / n1) + (var2 / n2)) ^ 0.5
            
            
            If counter > 0 Then
                outputarray(counter) = t
            End If
            output.Add CDbl(gene(0)), t
        End If
    End If
    
    counter = counter + 1
    
Loop
Close #1

For i = 0 To UBound(genescore)
genescore(i) = outputarray(i + 1)
genenumber(i) = i + 1

'List1.AddItem genenumber(i)
'List1.AddItem genescore(i)
Next i

sortout = sortrnd(genenumber, genescore)

For i = 0 To 1999
'List1.AddItem genenumber(i)
'List1.AddItem genescore(i)
List1.AddItem "Gene seq  " & CStr(genenumber(i)) & vbTab & "Score =   " & CStr(genescore(i))
Next i

End Sub
