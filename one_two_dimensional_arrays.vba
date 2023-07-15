Private Sub CommandButton1_Click()
Dim A(), B() As Long
Dim n, i As Long
n = InputBox("")
ReDim A(n)
For i = 1 To n
A(i) = InputBox("dd" & i & "dd")
Cells(i, 1) = A(i)
If A(i) > Abs(10) Then
Cells(i, 2) = A(i)
End If
Next i
End Sub
'_________________________________________

Sub var_18_3()
Dim A(), B() As Double
Dim i, n, l, h, nom, z As Integer
n = Val(InputBox("Кількість елементів масиву"))
M = InputBox("Введіть M")
ReDim A(n)
nom = InputBox("1 - вручну;  інше число - автоматично")
If nom = 1 Then
For i = 1 To n
A(i) = InputBox("Вводимо елемент № " & i)
Cells(i, 1) = A(i)
Next i
Else
h = Val(InputBox("h="))
l = Val(InputBox("l="))
For i = 1 To n
A(i) = Int((h - l + 1) * Rnd + l)
Cells(i, 1) = A(i)
Next i
End If
s = 0
For i = 1 To n
    If A(i) < M Then
    s = s + A(i)
    End If
Next i
MsgBox ("Сума елементів=" + Str(s))
For j = 1 To n
    For i = 1 To n - 1
      If Abs(A(i)) < Abs(A(i + 1)) Then
      tmp = Abs(A(i))
      A(i) = Abs(A(i + 1))
      A(i + 1) = tmp
      End If
    Next i
Next j
MsgBox ("Впорядкований масив в стовпчику 4")
For i = 1 To n
    Cells(i, 4) = A(i)
Next i
End Sub
'_____________________________________________
Sub sevenone()
Dim a(), x, y, buf As Long
Dim n, i As Long
n = InputBox("Введіть розмірність масиву")
ReDim a(n, n)
For i = 1 To n
For j = 1 To n
a(i, j) = InputBox("Введіть (" & " i " & "," & " j" & ")елемент")
Next j
Next i
For i = 1 To n
For j = 1 To n
Cells(i, j) = a(j, i)
Next j
Next i
MsgBox ("Масив виведено")
End Sub
'________________________________________
Sub laborseventwo()
Dim a() As Integer
Dim n, i As Long
n = InputBox("Введіть розмірність матриці")
ReDim a(n, n)
For i = 1 To n - 1
For j = 2 To n
a(i, j) = j - i
If a(i, j) < 0 Then
a(i, j) = 0
End If
Cells(i, j) = a(i, j)
Next j
Next i
MsgBox ("Матриця виведена на екран")
End Sub
'__________________________________________
Sub var_16_3()
Dim a(), min, c, sum, tmp As Double
Dim n, i, j As Double
n = InputBox("Введіть кількість елементів")
ReDim a(n)
ReDim B(n)
For i = 1 To n
a(i) = InputBox("Задайте елемент масиву A(" + Str(i) + ")=")
Cells(i, 1) = a(i)
Next i
MsgBox ("Масив виведено у перший стовпчик додатку")
minn = a(1)
For i = 1 To n
    If a(i) < minn Then
    minn = a(i)
    End If
Next i
c = 0
dob = 1
For i = 1 To n
    If a(i) = minn Then
    c = c + 1
    dob = dob * i
    End If
    Next i
For j = 1 To n
    For i = 1 To n - 1
      If Abs(a(i)) > Abs(a(i + 1)) Then
      tmp = Abs(a(i))
      a(i) = Abs(a(i + 1))
      a(i + 1) = tmp
      End If
    Next i
Next j
MsgBox ("Кількість входжень мінімального елементу в стовпчику 2")
Cells(1, 2) = c
MsgBox ("Добуток індексів входжень мінімальних елементів в стовпчику 3")
Cells(1, 3) = dob
MsgBox ("Впорядкований масив в стовпчику 4")
For i = 1 To n
    Cells(i, 4) = a(i)
Next i
End Sub
'________________________________
Sub laborsix()
Dim a(), dob_1, dob_2 As Long
Dim n, i As Long
n = InputBox("Введіть кількість елементів масиву")
ReDim a(n)
ReDim B(n)
For i = 1 To n
a(i) = InputBox("Задайте елемент масиву A(" + Str(i) + ")=")
Cells(i, 1) = a(i)
Next i
MsgBox ("Масив виведено у перший стовпчик додатку")
For i = 1 To n
If i <= n / 2 Then
  dob_1 = (Cells(i, 1)) * 2
  Cells(i, 2) = dob_1
Else
  dob_2 = (Cells(i, 1)) * 3
  Cells(i, 2) = dob_2
End If
Next i
End Sub
'___________________________
Sub labor5()
Dim i, sum   As Long
sum = 0
For i = 11 To 99 Step 2
sum = (sum + (i ^ 2))
Next i
MsgBox ("Сума=" & sum)
End Sub
'________________________________
Sub mas()
Dim a, B, h, min, max, sum, dob As Single
Dim i, j As Integer
a = Val(InputBox("Введіть верхню межу a"))
B = Val(InputBox("Введіть нижню межу b"))
h = Val(InputBox("Введіть крок h"))
x = a
i = 1
Do While (x <= B)
y = Abs(Tan(x) ^ 3)
Cells(i, 1) = x
Cells(i, 2) = y
i = i + 1
x = x + h
Loop
 max = Cells(2, 2)
For i = 2 To i - 1
    If Cells(i, 2) > max Then
        max = Cells(i, 2)
        End If
Next i
min = Cells(2, 2)
For i = 2 To i - 1
If Cells(i, 2) < min Then
min = Cells(i, 2)
End If
Next i
sum = 0
dob = 1
For i = 1 To i - 1
    If Cells(i, 2) >= 0 Then
        sum = sum + Cells(i, 2)
        Else
        dob = dob * Cells(i, 2)
    End If
Next i
Cells(1, 3) = "Min"
Cells(2, 3) = min
Cells(1, 4) = "Max"
Cells(2, 4) = max
Cells(1, 5) = "Sum"
Cells(2, 5) = sum
Cells(1, 6) = "Dob"
Cells(2, 6) = dob
End Sub

'_______________________________________
Public Sub завдання_1()
Dim A() As Integer
Dim n, i, j As Byte
n = InputBox("введіть кількість елементів масиву")
ReDim A(n, n)
For i = 1 To n
For j = 1 To n
    If i = j Or (i + j) = n + 1 Then
    Cells(i, j) = 1
    Else
    Cells(i, j) = 0
    End If
Next j
Next i
End Sub
'___________________________________________
Public Sub завд2()
    Dim A(), i, j As Integer
    Dim k As String
k = InputBox("введіть кількість елементів масиву")
ReDim A(k, k)
For i = 1 To k
    For j = 1 To k
        If (j = 1 Or j = k) And (i = 2 Or i = k - 1) Or (j = 2 Or j = k - 1) And (i = k Or i = 1) Then
        Cells(i, j) = 0
        Else
        Cells(i, j) = 1
        End If
    Next j
Next i
End Sub
'_____________________________________________
Public Sub завд3()
Dim A(), i, j As Integer
    Dim n As String
n = InputBox("введіть кількість елементів масиву")
ReDim A(n, n)
For i = 1 To n
    For j = 1 To n
        If (j = 1 Or j = n) And (i > 1 And i < n) Or (j = 2 Or j = n - 1) And (i > 2 And i < n - 1) Or (j = 3 Or j = n - 2) And (i > 3 And i < n - 2) Then
        Cells(i, j) = 0
        Else
        Cells(i, j) = 1
        End If
    Next j
    Next i
End Sub
'____________________________________
private завд4()
Dim A(), i, j, n As Integer
n = InputBox("Введіть розмірність матриці")

ReDim A(n, n)

For i = 1 To n
    For j = 1 To i
        A(j - 1, i - j) = i 'над діагоналлю
        A(n - i + j - 1, n - j) = i 'під діагоналлю
    Next j
Next i

    Dim res As String
    res = ""
    For i = 0 To n - 1
    For j = 0 To n - 1
    res = res & "  " & A(i, j)
    Next j
    res = res & vbCrLf
    Next i
    MsgBox (res)
End Sub
