Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

<Serializable()> Public Class clsRenter

    Protected name As String
    Protected ssNumber As String
    Protected address1 As String
    Protected address2 As String
    Protected city As String
    Protected state As String
    Protected zip As String
    Protected phone As String

    Public Sub New(ByVal name As String, ByVal ssnumber As String, ByVal address1 As String, ByVal address2 As String, ByVal city As String, ByVal state As String, ByVal zip As String, ByVal phone As String)
        Me.address1 = address1
        Me.address2 = address2
        Me.city = city
        Me.name = name
        Me.phone = phone
        Me.ssNumber = ssnumber
        Me.state = state
        Me.zip = zip
    End Sub
    Public Sub New(ByVal renter As clsRenter)
        Me.address1 = renter.getAddress1
        Me.address2 = renter.getAddress2
        Me.city = renter.getCity
        Me.name = renter.getName
        Me.phone = renter.getPhone
        Me.ssNumber = renter.getSsNumber
        Me.state = renter.getState
        Me.zip = renter.getZip
    End Sub
    Public Sub New(ByVal v() As String)
        Me.name = v(0)
        Me.ssNumber = v(1)
        Me.address1 = v(2)
        Me.address2 = v(3)
        Me.city = v(4)
        Me.state = v(5)
        Me.zip = v(6)
        Me.phone = v(7)
    End Sub

    Public Function getAddress1() As String
        Return address1
    End Function
    Public Function getAddress2() As String
        Return address2
    End Function
    Public Function getCity() As String
        Return city
    End Function
    Public Function getName() As String
        Return name
    End Function
    Public Function getPhone() As String
        Return phone
    End Function
    Public Function getSsNumber() As String
        Return ssNumber
    End Function
    Public Function getState() As String
        Return state
    End Function
    Public Function getZip() As String
        Return zip
    End Function
End Class

<Serializable()> Public Class clsTenant

    Protected renter() As clsRenter
    Protected apartment As String
    Protected priorApartment As String
    <Serializable()> Protected Structure termStruc
        Dim years As String
        Dim months As String
        Dim days As String
    End Structure
    Protected term As termStruc
    Protected leaseStart As Date
    Protected leaseEnd As Date
    Protected legalRent As Decimal
    Protected paidRent As Decimal
    Protected promptPayDisc As Decimal
    Protected security As Decimal
    Protected priorSecurity As Decimal
    ' Document fields
    Protected buildingId As Integer
    Protected printDocId As Integer
    Protected riders() As Integer
    Protected leaseId As Integer
    Protected savedOn As Date

    Public Sub New(ByVal renters() As clsRenter, ByVal apartment As String, ByVal years As Integer, ByVal months As Integer, ByVal days As Integer, ByVal leaseStart As Date, ByVal leaseEnd As Date, _
    ByVal legalRent As Decimal, ByVal paidRent As Decimal, ByVal security As Decimal, ByVal priorapartment As String, ByVal priorsecurity As Decimal, ByVal promptpaydisc As Decimal, _
    ByVal buildingId As Integer, ByVal printDocId As Integer, ByVal riders() As Integer, ByVal LeaseId As Integer)

        Dim i As Int16
        For i = 0 To renters.Length - 1
            If Not IsNothing(renters(i)) Then
                ReDim Preserve renter(i)
                renter(i) = New clsRenter(renters(i))
            End If
        Next
        Me.apartment = apartment
        Me.priorApartment = priorapartment
        Me.leaseEnd = leaseEnd
        Me.leaseStart = leaseStart
        Me.legalRent = legalRent
        Me.paidRent = paidRent
        Me.promptPayDisc = promptpaydisc
        Me.security = security
        Me.priorSecurity = priorsecurity
        term.years = Format(years, "##")
        term.months = Format(months, "##")
        term.days = Format(days, "##")
        Me.printDocId = printDocId
        Me.buildingId = buildingId
        Me.riders = riders
        Me.leaseId = LeaseId

    End Sub
    Public Function getField(ByVal fldName As String, Optional ByVal n As Integer = -1) As Object

        If n < -1 Then Return Nothing
        If n > -1 Then
            If IsNothing(renter) Then
                Return Nothing
            End If
            If n > renter.Length Then
                Return Nothing
            End If
        End If
        Select Case LCase(Mid(fldName, 1, 3))
            Case "bld"
                Select Case LCase(Mid(fldName, 4))
                    Case "name"
                        Return printBuilding.getName()
                    Case "address1"
                        Return printBuilding.getAddress1
                    Case "address2"
                        Return printBuilding.getAddress2
                    Case "city"
                        Return printBuilding.getCity
                    Case "state"
                        Return printBuilding.getState
                    Case "zip"
                        Return printBuilding.getZip
                    Case "contact"
                        Return printBuilding.getContact
                    Case Else
                        Return Nothing
                End Select
            Case Else
                Select Case LCase(Mid(fldName, 1, 5))
                    Case "owner"
                        Select Case LCase(Mid(fldName, 6))
                            Case "name"
                                Return printBuilding.getOwneName
                            Case "address1"
                                Return printBuilding.getOwnerAddress1
                            Case "address2"
                                Return printBuilding.getOwnerAddress2
                            Case "address3"
                                Return printBuilding.getOwnerAddress3
                            Case "telephone"
                                Return printBuilding.getTelephone
                            Case "fax"
                                Return printBuilding.getFax
                            Case "bank"
                                Return printBuilding.getOwnerBank
                            Case Else
                                Return Nothing
                        End Select
                    Case Else
                        Select Case LCase(fldName)
                            Case "apartment"
                                Return apartment
                            Case "priorapartment"
                                Return priorApartment
                            Case "leaseend"
                                Return IIf(leaseEnd < "#2000-01-01#", "", leaseEnd)
                            Case "leaseendlong"
                                Return IIf(leaseEnd < "#2000-01-01#", "", Format(leaseEnd, "MMMM d, yyyy"))
                            Case "leasestart"
                                Return IIf(leaseStart < "#2000-01-01#", "", leaseStart)
                            Case "leasestartlong"
                                Return IIf(leaseStart < "#2000-01-01#", "", Format(leaseStart, "MMMM d, yyyy"))
                            Case "legalrent"
                                Return legalRent
                            Case "paidrent"
                                Return paidRent
                            Case "promptpaydisc"
                                Return promptPayDisc
                            Case "security"
                                Return security
                            Case "priorsecurity"
                                Return priorSecurity
                            Case "term"
                                Return CStr(IIf(Val(term.years) > 0, term.years & CStr(IIf(Val(term.years) > 1, "s ", " ")), "")) _
                                & CStr(IIf(Val(term.months) > 0, term.months & CStr(IIf(Val(term.months) > 1, "s ", " ")), "")) _
                                & CStr(IIf(Val(term.days) > 0, term.days & CStr(IIf(Val(term.days) > 1, "s ", " ")), ""))
                            Case "termyears"
                                Return term.years
                            Case "termmonths"
                                Return term.months
                            Case "termdays"
                                Return term.days
                            Case "renternumber"
                                Return renter(n - 1)
                            Case "rentername"
                                Return renter(n - 1).getName
                            Case "renteraddress1"
                                Return renter(n - 1).getAddress1
                            Case "renteraddress2"
                                Return renter(n - 1).getAddress2
                            Case "rentercity"
                                Return renter(n - 1).getCity
                            Case "renterstate"
                                Return renter(n - 1).getState
                            Case "renterzip"
                                Return renter(n - 1).getZip
                            Case "renterss"
                                Return renter(n - 1).getSsNumber
                            Case "renterphone"
                                Return renter(n - 1).getPhone
                            Case Else
                                Return Nothing
                        End Select
                End Select
        End Select

        Return Nothing

    End Function

    Public Function serialize() As Boolean
        savedOn = Now()
        Dim fl As FileStream = Nothing
        Dim filename As String = renter(0).getName
        If IsNothing(filename) Then
            Return False
        End If
        If Trim(filename) = "" Then
            Return False
        End If
        fl = File.OpenWrite(datPath & filename & ".ser")
        Dim bf As BinaryFormatter = New BinaryFormatter
        bf.Serialize(fl, Me)
        fl.Close()
        Return True
    End Function

    Public Shared Function deSerialize(ByVal filename As String) As clsTenant

        Dim fl As FileStream = Nothing
        Try
            fl = File.OpenRead(datPath & filename)
            Dim bf As BinaryFormatter = New BinaryFormatter
            Dim obj As Object = bf.Deserialize(fl)
            fl.Close()
            If Not IsNothing(obj) Then
                Return DirectCast(obj, clsTenant)
            Else
                Return Nothing
            End If
        Catch
            Return Nothing
        End Try

    End Function

    Public Function getNumberRenters() As Integer
        Return renter.Length
    End Function
    Public Function getRenters() As clsRenter()
        Return renter
    End Function
    Public Function getRenter(ByVal i As Integer) As clsRenter
        Return renter(i)
    End Function
    Public Function getApartment() As String
        Return apartment
    End Function
    Public Function getPriorApartment() As String
        Return priorApartment
    End Function
    Public Function getLeaseEnd() As Date
        Return leaseEnd
    End Function
    Public Function getLeaseStart() As Date
        Return leaseStart
    End Function
    Public Function getLegalRent() As Decimal
        Return legalRent
    End Function
    Public Function getPaidRent() As Decimal
        Return paidRent
    End Function
    Public Function getPromptPayDisc() As Decimal
        Return promptPayDisc
    End Function
    Public Function getSecurity() As Decimal
        Return security
    End Function
    Public Function getPriorSecurity() As Decimal
        Return priorSecurity
    End Function
    Public Function getTermYears() As String
        Return term.years
    End Function
    Public Function getTermMonths() As String
        Return term.months
    End Function
    Public Function getTermDays() As String
        Return term.days
    End Function
    Public Function getTerm() As String

        Return CStr(IIf(Val(term.years) > 0, term.years & " year" & CStr(IIf(Val(term.years) > 1, "s ", " ")), "")) _
                & CStr(IIf(Val(term.months) > 0, term.months & " month" & CStr(IIf(Val(term.months) > 1, "s ", " ")), "")) _
                & CStr(IIf(Val(term.days) > 0, term.days & " day" & CStr(IIf(Val(term.days) > 1, "s ", " ")), ""))

    End Function
    Public Function getBuildingId() As Integer
        Return buildingId
    End Function
    Public Function getPrintDocId() As Integer
        Return printDocId
    End Function
    Public Function getSavedOn() As Date
        Return savedOn
    End Function
    Public Function getRiders() As Integer()
        Return riders
    End Function
    Public Function getLeaseId() As Integer
        Return leaseId
    End Function
End Class
