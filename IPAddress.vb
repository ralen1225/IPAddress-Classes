Option Explicit On
Option Strict On

Public Class IPAddress
  ' IP and mask values
  Private _ip As String
  Private _mask As UInteger

  ' IP property set and get to insulate private variable
  Public Property IP As String
    Get
      Return _ip
    End Get
    Set(value As String)
      _ip = value
    End Set
  End Property

  ' Mask property set and get to insulate private variable
  Public Property Mask As UInteger
    Get
      Return _mask
    End Get
    Set(value As UInteger)
      _mask = value
    End Set
  End Property

  ' Inputs string and integer
  Public Sub New(ByVal IPAdd As String, ByVal IPMask As UInteger)
    _ip = IPAdd
    _mask = IPMask
  End Sub

  ' Inputs string and string
  Public Sub New(ByVal IPAdd As String, ByVal IPMask As String)
    _ip = IPAdd
    UInteger.TryParse(IPMask, _mask)
  End Sub

  ' Input to duplicate 
  Public Sub New(ByRef IPAdd As IPAddress)
    _ip = IPAdd.IP
    _mask = IPAdd.Mask
  End Sub

  ' Numeric formatted IP address to 32 bits
  Public ReadOnly Property Numeric As UInteger
    Get
      ' Accumulator
      Dim intReturn As UInteger = 0
      ' Loop for 4 octets
      For i = 1 To 4
        ' Shift and insert into number
        intReturn = intReturn Or (Octet(i) << (8 * (4 - i)))
      Next
      ' Return result
      Return intReturn
    End Get
  End Property

  ' Calculate subnet IP address
  Public ReadOnly Property Subnet As String
    Get
      ' Return new subnet IP address
      Return CalcIP(Numeric And NetMask)
    End Get
  End Property

  ' Mask Converts XX netmask to 32 bits
  Public ReadOnly Property NetMask As UInteger
    Get
      Dim intMask As UInteger = 0
      ' Loop
      For i = 1 To 32
        ' Multiply mask by 2 (Shift left 1)
        intMask <<= 1
        ' Add 1 if appropriate
        If i <= Mask Then intMask += 1UI
      Next
      Return intMask
    End Get
  End Property

  Public ReadOnly Property Octet(selected As Integer) As UInteger
    Get
      ' This function will not assume that IPAddress is formatted properly X.X.X.X
      ' Return value
      Dim intReturn As UInteger
      ' Octet string array
      Dim strOctets() As String = IP.Split("."c)
      ' Make sure there are at least 4 elements
      Do Until strOctets.Length >= 4
        ' Add an element
        ReDim strOctets(strOctets.Length)
        ' Zero new element
        strOctets(strOctets.GetUpperBound(0)) = "0"
      Loop
      ' Select Octet and return it
      UInteger.TryParse(strOctets(selected - 1), intReturn)
      Return intReturn
    End Get
  End Property

  ' ToBinary converts 32 bits to binary string
  Public ReadOnly Property ToBinary(Optional ConvertMask As Boolean = False) As String
    Get
      ' Output string
      Dim strOut As String = ""
      ' Which number are we working on?
      Dim intNumber As UInteger = If(ConvertMask, NetMask(), Numeric())
      ' Loop for 32 bits
      For i = 0 To 31
        ' Add 1 for on bit, 0 for off bit
        strOut += If((intNumber And (1UI << (31 - i))) <> 0, "1", "0")
      Next
      ' Return result
      Return strOut
    End Get
  End Property

  ' Test to see if this is a valid IP address
  Public ReadOnly Property Valid As Boolean
    Get
      ' Valid character list
      Dim strValidChars As String = "0123456789."
      ' Check each character
      For i = 0 To IP.Length - 1
        If Not strValidChars.Contains(IP.ToUpper.Substring(i, 1)) Then Return False
      Next
      ' Split string into octets
      Dim strOctets() As String = IP.Split("."c)
      ' If not 4 octets, it fails
      If strOctets.Length <> 4 Then Return False
      ' Check each octet
      For Each strOctet In strOctets
        ' Check length - 0-length strings will not convert to UInt
        If strOctet.Length < 1 Then Return False
        ' Check value - guaranteed to be digits only at this point
        If CUInt(strOctet) > 255 Then Return False
      Next
      ' Netmask out of range
      If _mask > 32 Then Return False
      ' Passed all tests
      Return True
    End Get
  End Property

  ' Calculate Summary IP address and keep it
  Public Sub Summarize(ByRef IPInput As IPAddress)
    ' If values match, don't do anything
    If Not IPInput.Numeric = Numeric Then
      ' Locate first difference
      Dim IPXOr As New IPAddress(CalcIP(Not (Numeric Xor IPInput.Numeric)), 0)
      ' Set new mask
      _mask = CUInt(IPXOr.ToBinary.IndexOf("0") - 1)
      ' Calculate new IP
      _ip = CalcIP(NetMask And Numeric And IPInput.Numeric)
    End If
  End Sub

  ' CalcIP Converts 32 bits to string IP address
  Public Function CalcIP(IPNumber As UInteger) As String
    ' Empty string
    Dim strIPAddr As String = ""
    ' Generate 4 octets
    For i = 0 To 3
      ' After the first one, add a dot
      If i > 0 Then strIPAddr += "."
      ' Isolate current octet
      strIPAddr += ((IPNumber >> ((3 - i) * 8)) And &HFFUI).ToString
    Next
    ' Return IP string
    Return strIPAddr
  End Function
End Class
