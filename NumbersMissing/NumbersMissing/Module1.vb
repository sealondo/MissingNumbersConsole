Module Module1

    Sub Main()

        'Declaración de variables
        Dim sLonListA, sListA, sLonListB, sListB, sResultado As String

        'Recibir valores de entrada
        Console.WriteLine("Introduzca el tamaño de la lista A:")
        sLonListA = Console.ReadLine
        Console.WriteLine("Introduzca la lista A. Cada valor debe ir separado por un espacio:")
        sListA = Console.ReadLine
        Console.WriteLine("Introduzca el tamaño de la lista B:")
        sLonListB = Console.ReadLine
        Console.WriteLine("Introduzca la lista B. Cada valor debe ir separado por un espacio:")
        sListB = Console.ReadLine

        'Escribir el resultado
        Console.WriteLine("")
        Console.WriteLine("RESULTADO:")
        sResultado = LeerValores(sLonListA, sListA, sLonListB, sListB)
        Console.WriteLine(sResultado)
        Console.WriteLine("Presione cualquier tecla para finalizar.")
        Console.ReadKey()

    End Sub

    Function LeerValores(ByVal nLonListA As Integer, ByVal sListA As String, ByVal nLonListB As Integer, ByVal sListB As String)
        'Declaración de variables
        Dim aListA(), aListB() As String
        Dim nEnListA, nEnListB, j As Integer
        Dim sSalida As String

        'Inicialización de variables
        sSalida = ""

        'Construir arrays
        aListA = sListA.Split(" ")
        aListB = sListB.Split(" ")

        'Ordenar los arrays
        OrdenarArrays(aListA)
        OrdenarArrays(aListB)

        'Validar entradas
        If nLonListA < 1 Then
            Return "El tamaño de la lista A es menor a 1."
            Exit Function
        ElseIf nLonListA > 200000 Then
            Return "El tamaño de la lista A es mayor a 200000."
            Exit Function
        ElseIf nLonListB < 1 Then
            Return "El tamaño de la lista B es menor a 1."
            Exit Function
        ElseIf nLonListB > 200000 Then
            Return "El tamaño de la lista B es mayor a 200000"
        End If

        'Validar tamaño de A y B
        If nLonListA > nLonListB Then
            Return "El tamaño de la lista A es mayor al tamaño de la lista B"
            Exit Function
        End If

        'Validar valores de la lista B
        For Each b In aListB
            If CInt(b) > 10000 Then
                Return "El valor " & b & " de la lista B es mayor a 10000."
                Exit Function
            ElseIf CInt(b) < 1 Then
                Return "El valor " & b & " de la lista B es menor a 1."
                Exit Function
            End If
        Next

        'Validar diferencia entre valor máximo y valor mínimo
        If CInt(aListB(aListB.Length - 1)) - CInt(aListB(0)) > 100 Then
            Return "La diferencia entre el valor máximo: " & aListB(aListB.Length - 1) &
                    " y el valor mínimo: " & aListB(0) & " de la lista B supera 100."
            Exit Function
        End If

        'Encontrar los números que faltan
        For i = 0 To aListA.Length - 1
            nEnListA = 1
            nEnListB = 0
            If i <> aListA.Length - 1 Then
                j = i
                While aListA(j) = aListA(j + 1)
                    nEnListA = nEnListA + 1
                    j = j + 1
                    If j = (aListA.Length - 1) Then
                        Exit While
                    End If
                End While
            End If

            'Buscar en la lista B el valor de la lista A
            For k = 0 To aListB.Length - 1
                If aListB(k) = aListA(i) Then
                    nEnListB = nEnListB + 1
                End If
            Next k

            'Comparar la cantidad de veces que aparecen en ambas listas
            If nEnListA < nEnListB Then
                sSalida = sSalida & CStr(aListA(i)) & " "
            End If

            'Se aumenta hasta la posición del siguiente número
            i = i + (nEnListA - 1)

        Next i

        'Escribir valores perdidos
        If sSalida <> "" Then
            Console.WriteLine("Los valores perdidos fueron:")
            Console.WriteLine(sSalida)
        Else
            Console.WriteLine("No se encontrarn valores perdidos.")
            Console.WriteLine(sSalida)
        End If


        Return ""
    End Function

    Sub OrdenarArrays(ByRef aList As String())
        Dim tmp As Integer
        For i = 0 To aList.Length - 2
            For z = i + 1 To aList.Length - 1
                If CInt(aList(i)) > CInt(aList(z)) Then
                    tmp = CInt(aList(i))
                    aList(i) = aList(z)
                    aList(z) = tmp
                End If
            Next
        Next
    End Sub
End Module