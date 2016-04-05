Sub BaryCentric()

Dim pt, xlength, ylength, zlength, i, j, k As Integer
Dim xyz(1 To 4, 1 To 3), xyzPrime(1 To 4, 1 To 3) As Double
Dim abc(1 To 3, 1 To 3) As Double
Dim vector(1 To 3, 1 To 3), v0(1 To 3), v1(1 To 3), v2(1 To 3) As Double
Dim quad(1 To 4, 1 To 4, 1 To 3), normal(1 To 4, 1 To 3), CenterPt(1 To 4, 1 To 3), h(1 To 4), Volume(1 To 4) As Double
Dim quadpt(1 To 4, 1 To 3) As Integer
Dim Zstep, z, t(1 To 3) As Double
Dim xpt, ypt, Xmax, Ymax, Xmin, Ymin, Xstep, Ystep As Double
Dim volumeT, VolumePercent(1 To 4), V(1 To 4) As Double



Zstep = 0.1
Xstep = 0.1
Ystep = 0.1

For pt = 1 To 4
    'populates current xyz data values
    xyz(pt, 1) = range("D" & pt + 1)
    xyz(pt, 2) = range("E" & pt + 1)
    xyz(pt, 3) = range("F" & pt + 1)
    
    'populates initial xyz data values for parametric equations
    xyzPrime(pt, 1) = range("D" & pt + 1)
    xyzPrime(pt, 2) = range("E" & pt + 1)
    xyzPrime(pt, 3) = range("F" & pt + 1)

Next pt


'designates the xyz values that will be used for each 3D quadrant
'quadrant 1
quadpt(1, 1) = 1
quadpt(1, 2) = 3
quadpt(1, 3) = 2

'quadrant 2
quadpt(2, 1) = 1
quadpt(2, 2) = 2
quadpt(2, 3) = 4
                    
'quadrant 3
quadpt(3, 1) = 1
quadpt(3, 2) = 4
quadpt(3, 3) = 3
                    
'quadrant 4
quadpt(4, 1) = 2
quadpt(4, 2) = 4
quadpt(4, 3) = 3
                    
'quad(quadrant,pt,xyz)
                    
For i = 1 To 4
    'quadrant base number
    For j = 1 To 3
        'select base point
        For k = 1 To 3
            'input x,y,z values to quadrant point
            quad(i, j, k) = xyzPrime(quadpt(i, j), k)
        Next k
    Next j
Next i

'find total volume
volumeT = determinate(quad(1, 1, 1), quad(1, 1, 2), quad(1, 1, 3), quad(1, 2, 1), quad(1, 2, 2), quad(1, 2, 3), quad(1, 3, 1), quad(1, 3, 2), quad(1, 3, 3)) * xyz(4, 3) / 3


'parametric equations have the form x = x0+ta, y=y0+tb, z= z0+tc
'the folowing loop calculates the a,b,c for each parametric
'line from the base to the point of the prism
For pt = 1 To 3
    
    For i = 1 To 3

    vector(pt, i) = xyz(4, i) - xyz(pt, i)
    'vector(line,x,y,z component) from initial vertice point to middle z point
    
    Next i

Next pt
'populated vector(line,xyz) with each lines vector components

'find normal vector to each quadrant plane
'n(quadrant, i,j,k vector components)

For i = 1 To 4
    'quadrant
    v0(1) = xyz(quadpt(i, 2), 1) - xyz(quadpt(i, 1), 1)
    v0(2) = xyz(quadpt(i, 2), 2) - xyz(quadpt(i, 1), 2)
    v0(3) = xyz(quadpt(i, 2), 3) - xyz(quadpt(i, 1), 3)
        
    v1(1) = xyz(quadpt(i, 3), 1) - xyz(quadpt(i, 1), 1)
    v1(2) = xyz(quadpt(i, 3), 2) - xyz(quadpt(i, 1), 2)
    v1(3) = xyz(quadpt(i, 3), 3) - xyz(quadpt(i, 1), 3)
        
    normal(i, 1) = v0(2) * v1(3) - v0(3) * v1(2)
    normal(i, 2) = v0(3) * v1(1) - v0(1) * v1(3)
    normal(i, 3) = v0(1) * v1(2) - v0(2) * v1(1)
Next i
                        
                        
                        


zlength = xyz(4, 3) / Zstep
'determine size of z iterations

z = 0
For i = 1 To zlength
    
    'creates a rectangle from the lowest to highest of x and y values
    Xmax = 0
    Ymax = 0
    Xmin = xyz(1, 1)
    Ymin = xyz(1, 2)
    
    For j = 1 To 3
    
        If xyz(j, 1) > Xmax Then
            Xmax = xyz(j, 1)
        End If
    
        If xyz(j, 2) > Ymax Then
            Ymax = xyz(j, 2)
        End If
    
        If xyz(j, 1) < Xmin Then
            Xmin = xyz(j, 1)
        End If
    
        If xyz(j, 2) < Ymin Then
            Ymin = xyz(j, 2)
        End If
        'find max/min values for current z height
        
    Next j
    
    
    'v0,v1,v2 are made up of the test point and the triangles vertices. Cross product
    'and dot product calculation will determine if the test point is within the area
    'of interest
    For j = 1 To 3
       v0(j) = xyz(2, j) - xyz(1, j)
    Next j
    
    For j = 1 To 3
        v1(j) = xyz(3, j) - xyz(1, j)
    Next j
    
    v2(3) = z - xyz(1, 3)
    'populate intriangle parameters
    

   
    'determine x,y iteration size at current z height
    xlength = Round((Xmax - Xmin) / Xstep, 0)
    ylength = Round((Ymax - Ymin) / Ystep, 0)
        
    
    'set x start
    x = Xmin
    
        
    For xpt = 0 To xlength
            
        'populate x components of intriangle test point
        v2(1) = x - xyz(1, 1)
            
            'set y start
            y = Ymin
            
            For ypt = 0 To ylength
                
                'populate y component of intriangle test point
                v2(2) = y - xyz(1, 2)
                
                'test to see if point is inside triangle
                If inTriangle(v0(1), v0(2), v0(3), v1(1), v1(2), v1(3), v2(1), v2(2), v2(3)) = True Then
                
                    'range("D5") = x
                    'range("E5") = y
                    'run some code
                    
                    
                    For k = 1 To 4
                        'populate point of interest quadrant x,y,z values
                        quad(k, 4, 1) = x
                        quad(k, 4, 2) = y
                        quad(k, 4, 3) = z
                
                        'Application.Wait (Now + 0.00001)
                        
                        'find center of each quadrant
                        'quad(quadrant,pt,xyz)
                    
                        For j = 1 To 3
                            'xyz value
                            CenterPt(k, j) = average(quad(k, 1, j), quad(k, 2, j), quad(k, 3, j))
                        Next j
                    
                        'find h from center to POI
                        h(k) = length(CenterPt(k, 1), CenterPt(k, 2), CenterPt(k, 3), x, y, z)
                    
                        'find volume of quadrant (base times height)
                        V(k) = determinate(quad(k, 1, 1), quad(k, 1, 2), quad(k, 1, 3), quad(k, 2, 1), quad(k, 2, 2), quad(k, 2, 3), quad(k, 3, 1), quad(k, 3, 2), quad(k, 3, 3)) * h(k) / 3
                    
                        'find percentages of volume
                        VolumePercent(k) = V(k) / volumeT
                        
                        'volumepercent(1) * values(pt4) + volumepercent(2)* values(pt3) +volumepercent(3)* values(pt2) + volumepercent(4) * values(pt1)
                        
                        
                    Next k
                    
                    
                    
                    
                End If
                
                y = y + Ystep
            Next ypt
            
            x = x + Xstep
       Next xpt
       
        
    z = z + Zstep

    'increase in z direction by one unit
    j = 1
    For j = 1 To 3
        t(j) = (z - xyzPrime(j, 3)) / vector(j, 3)
        'determine t for each vector

        xyz(j, 1) = xyzPrime(j, 1) + vector(j, 1) * t(j)
        xyz(j, 2) = xyzPrime(j, 2) + vector(j, 2) * t(j)
        xyz(j, 3) = xyzPrime(j, 3) + vector(j, 3) * t(j)
        'increase parametric equation by t
    
        range("D" & j + 1) = xyz(j, 1)
        range("E" & j + 1) = xyz(j, 2)
        range("F" & j + 1) = xyz(j, 3)
        'graphical confirmation
    Next j
    
    
    
Next i

For pt = 1 To 4
    
    range("D" & pt + 1) = xyzPrime(pt, 1)
    range("E" & pt + 1) = xyzPrime(pt, 2)
    range("F" & pt + 1) = xyzPrime(pt, 3)

Next pt




End Sub


Function length(x1, y1, z1, x2, y2, z2)
Dim l As Double

length = ((x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (z2 - z1) ^ 2) ^ 0.5



End Function
Function determinate(x1, y1, z1, x2, y2, z2, x3, y3, z3)
Dim a(1 To 2, 1 To 3) As Double
Dim b(1 To 3) As Double

a(1, 1) = x2 - x1
a(1, 2) = y2 - y1
a(1, 3) = z2 - z1
a(2, 1) = x3 - x1
a(2, 2) = y3 - y1
a(2, 3) = z3 - z1

b(1) = a(1, 2) * a(2, 3) - a(1, 3) * a(2, 2)
b(2) = a(1, 1) * a(2, 3) - a(1, 3) * a(2, 1)
b(3) = a(1, 1) * a(2, 2) - a(1, 2) * a(2, 1)

determinate = 0.5 * ((b(1) ^ 2 + b(2) ^ 2 + b(3) ^ 2) ^ 0.5)




End Function

Function dot(a1, a2, a3, b1, b2, b3)

dot = a1 * b1 + a2 * b2 + a3 * b3

End Function

Function inTriangle(v01, v02, v03, v11, v12, v13, v21, v22, v23)
Dim d00, d01, d02, d11, d12 As Double
Dim invDenom As Double
Dim u, V As Double


d00 = dot(v01, v02, v03, v01, v02, v03)
d01 = dot(v01, v02, v03, v11, v12, v13)
d02 = dot(v01, v02, v03, v21, v22, v23)
d11 = dot(v11, v12, v13, v11, v12, v13)
d12 = dot(v11, v12, v13, v21, v22, v23)

    
invDenom = 1 / (d00 * d11 - d01 * d01)

u = (d11 * d02 - d01 * d12) * invDenom
V = (d00 * d12 - d01 * d02) * invDenom

inTriangle = False

If u >= 0 Then
    If V >= 0 Then
        If (u + V) < 1 Then
            inTriangle = True
        End If
    End If
End If
End Function

Function average(a, b, c)

average = (a + b + c) / 3


End Function



