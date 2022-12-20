Attribute VB_Name = "Meta"
Public Data_GenerateCount As Single
Public Data_MateCount(1 To 56) As Single

Public Protect As Boolean

Public Version As String

Public Class As String
Public MateAmount As Byte
Public MaleAmount As Byte
Public FemaleAmount As Byte
Public UnknowGenderAmount As Byte

Public Name(1 To 56) As String
Public Gender(1 To 56) As String

Public WindowState As String
Public WindowLastState As String

Public Amount As Single
Public LastAmount As Single
Public Result(1 To 1000000) As Single

Public GenerateTime
Public Old_GenerateTime

Public ViewLastData As Boolean
