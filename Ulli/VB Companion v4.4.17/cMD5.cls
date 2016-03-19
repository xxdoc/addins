VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'MD5 message digest
'based on the original by Ronald L. Rivest
'
' Feb 21 2002 UMG
'
' Update Aug 28 2004 UMG
' Changed ACs from constants to generated
' Changed returned signature to 16 bytes rather than 16 chars
' Made last two summands in UnsigendAdd optional
'
' Bugfix Feb 24 2006
' Changed ACs indexing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' The MD5 algorithm produces a 128 bit digital fingerprint (signature) from a
' dataset of arbitrary length. For details see RFC 1321 (summarized below). This
' implementation is derived from the RSA Data Security, Inc. MD5 Message-Digest
' algorithm reference implementation (originally written in C).
'
' NOTES:
'     Network Working Group                                             R. Rivest
'     Request for Comments: 1321              MIT Laboratory for Computer Science
'                                                      and RSA Data Security, Inc.
'                                                                      April 1992
'
'                          The MD5 Message-Digest Algorithm
'                          ================================
'     Summary
'     -------
'
'       This document describes the MD5 message-digest algorithm. The
'       algorithm takes as input a message of arbitrary length and produces
'       as output a 128-bit "fingerprint" or "message digest" of the input.
'       It is conjectured that it is computationally infeasible to produce
'       two messages having the same message digest, or to produce any
'       message having a given prespecified target message digest. The MD5
'       algorithm is intended for digital signature applications, where a
'       large file must be "compressed" in a secure manner before being
'       encrypted with a private (secret) key under a public-key cryptosystem
'       such as RSA.
'
'       The MD5 algorithm is designed to be quite fast on 32-bit machines. In
'       addition, the MD5 algorithm does not require any large substitution
'       tables; the algorithm can be coded quite compactly.
'
'       The MD5 algorithm is an extension of the MD4 message-digest algorithm.
'       MD5 is slightly slower than MD4, but is more "conservative" in
'       design. MD5 was designed because it was felt that MD4 was perhaps
'       being adopted for use more quickly than justified by the existing
'       critical review; because MD4 was designed to be exceptionally fast,
'       it is "at the edge" in terms of risking successful cryptanalytic
'       attack. MD5 backs off a bit, giving up a little in speed for a much
'       greater likelihood of ultimate security. It incorporates some
'       suggestions made by various reviewers, and contains additional
'       optimizations. The MD5 algorithm is being placed in the public domain
'       for review and possible adoption as a standard.
'
'       RFC Author:
'       Ronald L.Rivest
'       Massachusetts Institute of Technology
'       Laboratory for Computer Science
'       NE43-324545 Technology Square
'       Cambridge, MA 02139-1986
'       Phone: (617) 253-5880
'       EMail: Rivest@theory.lcs.mit.edu
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
DefLng A-Z 'we're 32 bit!

Private Const IniPart1  As Long = &H67452301 'x 01 23 45 67
Private Const IniPart2  As Long = &HEFCDAB89 'x 89 ab cd ef
Private Const IniPart3  As Long = &H98BADCFE 'x fe dc ba 98
Private Const IniPart4  As Long = &H10325476 'x 76 54 32 10

Private Const MaxLong   As Long = 2 ^ 31 - 1
Private Const ACFact    As Double = 2 ^ 32

Private ACs(1 To 64)    As Long

'round 1
Private Const S11       As Long = 7     '+ 5
Private Const S12       As Long = 12    '+ 5
Private Const S13       As Long = 17    '+ 5
Private Const S14       As Long = 22

'round 2
Private Const S21       As Long = 5     '+ 4
Private Const S22       As Long = 9     '+ 5
Private Const S23       As Long = 14    '+ 6
Private Const S24       As Long = 20

'round 3
Private Const S31       As Long = 4     '+ 7
Private Const S32       As Long = 11    '+ 5
Private Const S33       As Long = 16    '+ 7
Private Const S34       As Long = 23

'round 4
Private Const S41       As Long = 6     '+ 4
Private Const S42       As Long = 10    '+ 5
Private Const S43       As Long = 15    '+ 6
Private Const S44       As Long = 21

'some bit masks
Private Const Top16Bits As Long = &HFFFF0000
Private Const Bot16Bits As Long = &HFFFF&
Private Const Bit16Only As Long = &H10000

'the chunk size
Private Const BlockSize As Long = 64

'structures for converting long to byte
Private Type tFourLongs
    Longs(1 To 4)       As Long
End Type
Private FourLongs       As tFourLongs

Private Type tSixteenBytes
    Bytes(1 To 16)      As Byte
End Type
Private SixteenBytes    As tSixteenBytes

'misc
Private i               As Long
Private j               As Long
Private k               As Long

Private Sub Class_Initialize()

    For i = LBound(ACs) To UBound(ACs) 'create ACs
        ACs(i) = ToUnsignedLong(CCur(Int(Abs(ACFact * Sin(i)))))
    Next i

End Sub

Private Sub DigestBlock(Data As String)

  'MD5 mangle, requires a string of length 64

  Dim X(0 To 15)
  Dim a, b, c, d

  'pack 4 bytes into 1 long -> 64 bytes into 16 unsigned longs = 512 bits

    For i = 1 To BlockSize Step 4
        X(i \ 4) = RotateLeft(RotateLeft(RotateLeft(Asc(Mid$(Data, i + 3, 1)), 8) Or Asc(Mid$(Data, i + 2, 1)), 8) Or Asc(Mid$(Data, i + 1, 1)), 8) Or Asc(Mid$(Data, i, 1))
    Next i

    With FourLongs
        k = 1

        'round 1 digest
        a = Round1(.Longs(1), .Longs(2), .Longs(3), .Longs(4), X(0), S11, ACs(k))
        d = Round1(.Longs(4), a, .Longs(2), .Longs(3), X(1), S12, ACs(k))
        c = Round1(.Longs(3), d, a, .Longs(2), X(2), S13, ACs(k))
        b = Round1(.Longs(2), c, d, a, X(3), S14, ACs(k))

        a = Round1(a, b, c, d, X(4), S11, ACs(k))
        d = Round1(d, a, b, c, X(5), S12, ACs(k))
        c = Round1(c, d, a, b, X(6), S13, ACs(k))
        b = Round1(b, c, d, a, X(7), S14, ACs(k))

        a = Round1(a, b, c, d, X(8), S11, ACs(k))
        d = Round1(d, a, b, c, X(9), S12, ACs(k))
        c = Round1(c, d, a, b, X(10), S13, ACs(k))
        b = Round1(b, c, d, a, X(11), S14, ACs(k))

        a = Round1(a, b, c, d, X(12), S11, ACs(k))
        d = Round1(d, a, b, c, X(13), S12, ACs(k))
        c = Round1(c, d, a, b, X(14), S13, ACs(k))
        b = Round1(b, c, d, a, X(15), S14, ACs(k))

        'round 2 digest
        a = Round2(a, b, c, d, X(1), S21, ACs(k))
        d = Round2(d, a, b, c, X(6), S22, ACs(k))
        c = Round2(c, d, a, b, X(11), S23, ACs(k))
        b = Round2(b, c, d, a, X(0), S24, ACs(k))

        a = Round2(a, b, c, d, X(5), S21, ACs(k))
        d = Round2(d, a, b, c, X(10), S22, ACs(k))
        c = Round2(c, d, a, b, X(15), S23, ACs(k))
        b = Round2(b, c, d, a, X(4), S24, ACs(k))

        a = Round2(a, b, c, d, X(9), S21, ACs(k))
        d = Round2(d, a, b, c, X(14), S22, ACs(k))
        c = Round2(c, d, a, b, X(3), S23, ACs(k))
        b = Round2(b, c, d, a, X(8), S24, ACs(k))

        a = Round2(a, b, c, d, X(13), S21, ACs(k))
        d = Round2(d, a, b, c, X(2), S22, ACs(k))
        c = Round2(c, d, a, b, X(7), S23, ACs(k))
        b = Round2(b, c, d, a, X(12), S24, ACs(k))

        'round 3 digest
        a = Round3(a, b, c, d, X(5), S31, ACs(k))
        d = Round3(d, a, b, c, X(8), S32, ACs(k))
        c = Round3(c, d, a, b, X(11), S33, ACs(k))
        b = Round3(b, c, d, a, X(14), S34, ACs(k))

        a = Round3(a, b, c, d, X(1), S31, ACs(k))
        d = Round3(d, a, b, c, X(4), S32, ACs(k))
        c = Round3(c, d, a, b, X(7), S33, ACs(k))
        b = Round3(b, c, d, a, X(10), S34, ACs(k))

        a = Round3(a, b, c, d, X(13), S31, ACs(k))
        d = Round3(d, a, b, c, X(0), S32, ACs(k))
        c = Round3(c, d, a, b, X(3), S33, ACs(k))
        b = Round3(b, c, d, a, X(6), S34, ACs(k))

        a = Round3(a, b, c, d, X(9), S31, ACs(k))
        d = Round3(d, a, b, c, X(12), S32, ACs(k))
        c = Round3(c, d, a, b, X(15), S33, ACs(k))
        b = Round3(b, c, d, a, X(2), S34, ACs(k))

        'round 4 digest
        a = Round4(a, b, c, d, X(0), S41, ACs(k))
        d = Round4(d, a, b, c, X(7), S42, ACs(k))
        c = Round4(c, d, a, b, X(14), S43, ACs(k))
        b = Round4(b, c, d, a, X(5), S44, ACs(k))

        a = Round4(a, b, c, d, X(12), S41, ACs(k))
        d = Round4(d, a, b, c, X(3), S42, ACs(k))
        c = Round4(c, d, a, b, X(10), S43, ACs(k))
        b = Round4(b, c, d, a, X(1), S44, ACs(k))

        a = Round4(a, b, c, d, X(8), S41, ACs(k))
        d = Round4(d, a, b, c, X(15), S42, ACs(k))
        c = Round4(c, d, a, b, X(6), S43, ACs(k))
        b = Round4(b, c, d, a, X(13), S44, ACs(k))

        a = Round4(a, b, c, d, X(4), S41, ACs(k))
        d = Round4(d, a, b, c, X(11), S42, ACs(k))
        c = Round4(c, d, a, b, X(2), S43, ACs(k))
        b = Round4(b, c, d, a, X(9), S44, ACs(k))

        .Longs(1) = UnsignedAdd(.Longs(1), a)
        .Longs(2) = UnsignedAdd(.Longs(2), b)
        .Longs(3) = UnsignedAdd(.Longs(3), c)
        .Longs(4) = UnsignedAdd(.Longs(4), d)
    End With 'FOURLONGS

End Sub

Private Function RotateLeft(ByVal Value As Long, ByVal NumPlaces As Long) As Long

  'rotate a long numplaces to the left

  Dim CntPlaces, SpillOver

    For CntPlaces = 1 To NumPlaces Mod 32
        SpillOver = Value And &HC0000000
        Value = (Value And &H3FFFFFFF) * 2 Or ((SpillOver < 0) And 1) Or (CBool(SpillOver And &H40000000) And &H80000000)
    Next CntPlaces
    RotateLeft = Value

End Function

Private Function Round1(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal X As Long, ByVal s As Long, ByVal AC As Long) As Long

    Round1 = UnsignedAdd(RotateLeft(UnsignedAdd(a, (b And c) Or (d And Not b), X, AC), s), b)
    k = k + 1

End Function

Private Function Round2(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal X As Long, ByVal s As Long, ByVal AC As Long) As Long

    Round2 = UnsignedAdd(RotateLeft(UnsignedAdd(a, (b And d) Or (c And Not d), X, AC), s), b)
    k = k + 1

End Function

Private Function Round3(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal X As Long, ByVal s As Long, ByVal AC As Long) As Long

    Round3 = UnsignedAdd(RotateLeft(UnsignedAdd(a, b Xor c Xor d, X, AC), s), b)
    k = k + 1

End Function

Private Function Round4(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal X As Long, ByVal s As Long, ByVal AC As Long) As Long

    Round4 = UnsignedAdd(RotateLeft(UnsignedAdd(a, c Xor (b Or Not d), X, AC), s), b)
    k = k + 1

End Function

Public Function Signature(Data As String, Optional Reset As Boolean = True) As Byte()

  'digest a string and output the resulting signature as a string

  'this can be called several times in a row to sign a combination of more than one text;
  'first call must reset, subsequent calls may reset.

  'use the returned signatures to sign each indidual text or use the final signature to
  'sign the combination.

  'if you do not reset on subsequent calls then the intermediate signatures not only
  'depend on the digested text but also on previously digested texts.

  'or digest a password first (with reset) and then a text (without reset); the signature
  'will then reflect both the password and the text and will thus permit authenticication.

  Dim Asci      As Long
  Dim Padding   As String

    If Len(Data) Then
        If Reset Then 'initialize the algorithm
            With FourLongs
                .Longs(1) = ToUnsignedLong(IniPart1)
                .Longs(2) = ToUnsignedLong(IniPart2)
                .Longs(3) = ToUnsignedLong(IniPart3)
                .Longs(4) = ToUnsignedLong(IniPart4)
            End With 'FourLongs
        End If

        'create pseudo random padding, depending on data
        For j = 1 To Len(Data)
            Asci = Asci Xor Asc(Mid$(Data, j, 1))
        Next j
        Rnd -(Asci Xor Len(Data)) - 1 'bugfix: - 1 to prevent Rnd(0) if Asci and Len(Data) are equal
        Do
            Padding = Chr$(Int(Rnd * 255) + 1) & Padding
        Loop Until Len(Padding) = BlockSize - 1

        'digest the stuff in chunks
        For j = 1 To Len(Data) Step BlockSize
            DigestBlock Mid$(Data & Padding, j, BlockSize)
        Next j

        LSet SixteenBytes = FourLongs 'pack 4 longs into 16 bytes
        Signature = SixteenBytes.Bytes 'so the signature is 16 bytes = 128 bits chars long

    End If

End Function

Private Function ToUnsignedLong(Value As Currency) As Long

  'convert a currency datatype to an unsigned long

    If Value <= MaxLong Then
        ToUnsignedLong = Value
      Else 'NOT VALUE...
        ToUnsignedLong = Value - ACFact
    End If

End Function

Private Function UnsignedAdd(ByVal Summand1 As Long, ByVal Summand2 As Long, Optional ByVal Summand3 As Long = 0, Optional ByVal Summand4 As Long = 0) As Long

  'adds two or four unsigned numbers together, overflows are ignored

  Dim Low, High

    Low = (Summand1 And Bot16Bits) + (Summand2 And Bot16Bits) + (Summand3 And Bot16Bits) + (Summand4 And Bot16Bits)
    High = ((Summand1 And Top16Bits) \ Bit16Only + (Summand2 And Top16Bits) \ Bit16Only + (Summand3 And Top16Bits) \ Bit16Only + (Summand4 And Top16Bits) \ Bit16Only + Low \ Bit16Only) And Bot16Bits
    UnsignedAdd = ToUnsignedLong(CCur(High) * Bit16Only + (Low And Bot16Bits))

End Function

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 127  Code: 235  Total: 362 Lines
':) CommentOnly: 95 (26,2%)  Commented: 25 (6,9%)  Empty: 78 (21,5%)  Max Logic Depth: 4
