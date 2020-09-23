VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: NumberFormatInfo
'

''
' Provides numeric formatting information.
'
' @remarks<br>
' An instance of this class can be used to provide formatting for numeric data
' in places that allow for an <b>IFormatProvider</b> to be supplied. StringBuilder.AppendFormat
' is an example of when a <b>NumberFormatInfo</b> object can be used. Also, an
' instance of <b>NumberFormatInfo</b> can be used stand-alone to format numeric data
' by calling the <b>Format</b> method.
' <p>To format a number using <b>NumberFormatInfo</b>, call the <b>Format</b> method passing
' in the value to be formated.
' @see NumberFormatInfoStatic
' @see IFormatProvider
' @see ICloneable
' @include "..\..\Includes\NumberFormatInfo.txt"
Option Explicit
Implements cObject
Implements IFormatProvider
Implements ICloneable

Private Const PROP_POSITIVESIGN As String = "PositiveSign"
Private Const PROP_NEGATIVESIGN As String = "NegativeSign"
Private Const PROP_NUMBERDECIMALSEPARATOR As String = "NumberDecimalSeparator"
Private Const PROP_NUMBERDECIMALDIGITS As String = "NumberDecimalDigits"
Private Const PROP_NUMBERGROUPSIZES As String = "NumberGroupSizes"
Private Const PROP_NUMBERNEGATIVEPATTERN As String = "NumberNegativePattern"
Private Const PROP_NUMBERGROUPSEPARATOR As String = "NumberGroupSeparator"
Private Const PROP_CURRENCYDECIMALSEPARATOR As String = "CurrencyDecimalSeparator"
Private Const PROP_CURRENCYDECIMALDIGITS As String = "CurrencyDecimalDigits"
Private Const PROP_CURRENCYGROUPSIZES As String = "CurrencyGroupSizes"
Private Const PROP_CURRENCYNEGATIVEPATTERN As String = "CurrencyNegativePattern"
Private Const PROP_CURRENcYGROUPSEPARATOR As String = "CurrencyGroupSeparator"
Private Const PROP_CURRENCYSYMBOL As String = "CurrencySymbol"
Private Const PROP_CURRENCYPOSITIVEPATTERN As String = "CurrencyPositivePattern"
Private Const PROP_PERCENTDECIMALSEPARATOR As String = "PercentDecimalSeparator"
Private Const PROP_PERCENTDECIMALDIGITS As String = "PercentDecimalDigits"
Private Const PROP_PERCENTGROUPSIZES As String = "PercentGroupSizes"
Private Const PROP_PERCENTNEGATIVEPATTERN As String = "PercentNegativePattern"
Private Const PROP_PERCENTGROUPSEPARATOR As String = "PercentGroupSeparator"
Private Const PROP_PERCENTSYMBOL As String = "PercentSymbol"
Private Const PROP_PERCENTPOSITIVEPATTERN As String = "PercentPositivePattern"
Private Const PROP_PERMILLESYMBOL As String = "PerMilleSymbol"
Private Const PROP_NANSYMBOL As String = "NaNSymbol"
Private Const PROP_POSITIVEINFINITYSYMBOL As String = "PositiveInfinitySymbol"
Private Const PROP_NEGATIVEINFINITYSYMBOL As String = "NegativeInfinitySymbol"


Private Const INTEGER_PRECISION     As Long = 10
Private Const DOUBLE_PRECISION      As Long = 15
Private Const SINGLE_PRECISION      As Long = 7

Private Const CURRENCY_FORMAT       As Long = 67
Private Const DECIMAL_FORMAT        As Long = 68
Private Const EXPONENT_FORMAT       As Long = 69
Private Const FIXED_FORMAT          As Long = 70
Private Const GENERAL_FORMAT        As Long = 71
Private Const NUMBER_FORMAT         As Long = 78
Private Const PERCENT_FORMAT        As Long = 80
Private Const HEX_FORMAT            As Long = 88

Private Const CHAR_0                As Long = 48
Private Const CHAR_1                As Long = 49
Private Const CHAR_5                As Long = 53
Private Const CHAR_9                As Long = 57
Private Const CHAR_POUND            As Long = 35
Private Const CHAR_MINUS            As Long = 45
Private Const CHAR_DOLLAR           As Long = 36
Private Const CHAR_PERCENT          As Long = 37

Private Const SCALE_NAN             As Long = &H80000000
Private Const SCALE_INF             As Long = &H7FFFFFFF


''
' Output patterns for negative numbers.
'
' @param [(n)] Negative value inside perenthesis. Negative symbol is not included.
' @param [-n] Negative symbol before the value.
' @param [- n] Negative symbol before the value, separated by a space.
' @param [n-] Negative symbol after the value.
' @param [n -] Negative symbol after the value, separated by a space.
Public Enum NumberNegativePatterns
    [(n)] = 0
    [-n] = 1
    [- n] = 2
    [n-] = 3
    [n -] = 4
End Enum

''
' Output patterns for negative currency values. The <b>$</b> is the currency
' symbol, then <b>-</b> is the negative symbol and <b>n</b> is the number.
'
' @param [($n)]
' @param [-$n]
' @param [$-n]
' @param [$n-]
' @param [(n$)]
' @param [-n$]
' @param [n-$]
' @param [n$-]
' @param [-n $]
' @param [-$ n]
' @param [n $-]
' @param [$ n-]
' @param [$ -n]
' @param [n- $]
' @param [($ n)]
' @param [(n $)]
Public Enum CurrencyNegativePatterns
    [($n)] = 0
    [-$n] = 1
    [$-n] = 2
    [$n-] = 3
    [(n$)] = 4
    [-n$] = 5
    [n-$] = 6
    [n$-] = 7
    [-n $] = 8
    [-$ n] = 9
    [n $-] = 10
    [$ n-] = 11
    [$ -n] = 12
    [n- $] = 13
    [($ n)] = 14
    [(n $)] = 15
End Enum

''
' Output patterns for positive currency values. The <b>$</b> is the currency symbol and <b>n</b> is the number.
'
' @param [$n]
' @param [n$]
' @param [$ n]
' @param [n $]
Public Enum CurrencyPositivePatterns
    [$n] = 0
    [n$] = 1
    [$ n] = 2
    [n $] = 3
End Enum

''
' Output patterns for positive percentage values. The <b>#</b> is the number and the <b>%</b> is the percentage symbol.
'
' @param [# %]
' @param [#%]
' @param [%#]
Public Enum PercentPositivePatterns
    [# %] = 0
    [#%] = 1
    [%#] = 2
End Enum

''
' Output patterns for negative percentage values. The <b>-</b> is the negative symbol,
' <b>#</b> is the number and the <b>%</b> is the percentage symbol.
'
' @param [-# %]
' @param [-#%]
' @param [-%#]
Public Enum PercentNegativePatterns
    [-# %] = 0
    [-#%] = 1
    [-%#] = 2
End Enum


Private Type NumberType
    Scale As Long
    Precision As Long
    IsNegative As BOOL
    digits() As Byte
    DigitsSA As SafeArray1d
End Type
Private Type PropsType
    PositiveSign As String
    NegativeSign As String
    NumberDecimalSeparator As String
    NumberDecimalDigits As Long
    NumberGroupSizes() As Long
    NumberNegativePattern As NumberNegativePatterns
    NumberGroupSeparator As String
    CurrencyDecimalSeparator As String
    CurrencyDecimalDigits As Long
    CurrencyGroupSizes() As Long
    CurrencyNegativePattern As CurrencyNegativePatterns
    CurrencyGroupSeparator As String
    CurrencySymbol As String
    CurrencyPositivePattern As CurrencyPositivePatterns
    PercentDecimalSeparator As String
    PercentDecimalDigits As Long
    PercentGroupSizes() As Long
    PercentNegativePattern As PercentNegativePatterns
    PercentGroupSeparator As String
    PercentSymbol As String
    PercentPositivePattern As PercentPositivePatterns
    PerMilleSymbol As String
    NaNSymbol As String
    PositiveInfinitySymbol As String
    NegativeInfinitySymbol As String
    IsReadOnly As Boolean
End Type

' sets of string representation of the different output patterns.
Private mCurrencyPositivePatterns() As String
Private mCurrencyNegativePatterns() As String
Private mNumberNegativePatterns() As String
Private mPercentPositivePatterns() As String
Private mPercentNegativePatterns() As String

' utility variables
Private mBufferCapacity As Long
Private mBuffer As WordBuffer
Private mBufferPos As Long
Private mPtrNumberBuffer As Long
Private mNumber As NumberType
Private mFormatInfo As WordBuffer
Private mAppendStringBuffer As WordBuffer

' this is the set of NumberFormatInfo properties.
Private mProps As PropsType
Private mLoaded As Boolean


''
' Returns if this instance is read-only.
'
' @return Value indicating if this instance is read-only.
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mProps.IsReadOnly
End Property

' for internal use only.
Friend Property Let IsReadOnly(ByVal RHS As Boolean)
    mProps.IsReadOnly = RHS
End Property

''
' Returns a clone of this instance.
'
' @return A clone of the original instance.
' @remarks The <b>IsReadOnly</b> is passed to the clone, aswell.
Public Function Clone() As Object
    Dim ret As New NumberFormatInfo
    ret.CloneHelper mProps
    Set Clone = ret
End Function

''
' Returns a string representation of negative infinty.
'
' @return Negative infinity string.
' @remarks Negative infinity is achieved by dividing a negative
' number by zero. The default is "-Infinity".
Public Property Get NegativeInfinitySymbol() As String
    VerifyLoaded
    NegativeInfinitySymbol = mProps.NegativeInfinitySymbol
End Property

''
' Sets the negative infinity representation.
'
' @param RHS The new string that represents negative infinity.
' @remarks Negative infinity is achieved by dividing a negative
' number by zero. The default is "-Infinity".
Public Property Let NegativeInfinitySymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NegativeInfinitySymbol = RHS
End Property

''
' Returns a string representation of positive infinity.
'
' @return Positive infinity string.
' @remarks Positive infinity is achieved by dividing a positive
' number by zero. The default is "Infinity".
Public Property Get PositiveInfinitySymbol() As String
    VerifyLoaded
    PositiveInfinitySymbol = mProps.PositiveInfinitySymbol
End Property

''
' Sets the positive infinity representation.
'
' @param RHS The new string representation of positive infinity.
' @remarks Positive infinity is achieved by dividing a positive
' number by zero. The default is "Infinity".
Public Property Let PositiveInfinitySymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PositiveInfinitySymbol = RHS
End Property

''
' Returns the permille representation.
'
' @return Permille string representation.
' @remarks From <a href="http://mathworld.wolfram.com/Permil.html">MathWorld</a><br>
' The use of permille (a.k.a. parts per thousand) is a way of
' expressing ratios in terms of whole numbers. Given a ratio or fraction,
' it is converted to a permil-age by multiplying by 1000 and appending a "‰" (Unicode &h2030).
' For example, if an investment grows from a number P=13.00 to a number A=22.50,
' then A is 22.50/13.00=1.7308 times as much as P, or 1730.8‰.
Public Property Get PerMilleSymbol() As String
    VerifyLoaded
    PerMilleSymbol = mProps.PerMilleSymbol
End Property

''
' Sets the permille representation.
'
' @param RHS The new permille string representation.
' @remarks From From <a href="http://mathworld.wolfram.com/Permil.html">MathWorld</a><br>
' The use of permille (a.k.a. parts per thousand) is a way of
' expressing ratios in terms of whole numbers. Given a ratio or fraction,
' it is converted to a permil-age by multiplying by 1000 and appending a "‰" (Unicode &H2030).
' For example, if an investment grows from a number P=13.00 to a number A=22.50,
' then A is 22.50/13.00=1.7308 times as much as P, or 1730.8‰.
Public Property Let PerMilleSymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PerMilleSymbol = RHS
End Property

''
' Returns the string representation for Not-A-Number.
'
' @return The string representation.
' @remarks Not-A-Number is achieved by dividing zero by zero. The default
' NaN symbol is "NaN".
Public Property Get NaNSymbol() As String
    VerifyLoaded
    NaNSymbol = mProps.NaNSymbol
End Property

''
' Sets the string representation for Not-A-Number.
'
' @param RHS The new string NaN representation.
' @remarks Not-A-Number is achieved by dividing zero by zero. The default
' NaN symbol is "NaN".
Public Property Let NaNSymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NaNSymbol = RHS
End Property

''
' Returns the decimal symbol for percent formatted numbers.
'
' @return The symbol for decimal separation.
Public Property Get PercentDecimalSeparator() As String
    PercentDecimalSeparator = mProps.PercentDecimalSeparator
End Property

''
' Sets the decimal symbol for percent formatted numbers.
'
' @param RHS The new decimal symbol.
Public Property Let PercentDecimalSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentDecimalSeparator = RHS
End Property

''
' Returns the number of digits to be displayed after the
' decimal point for percent formatted numbers.
'
' @return The number of digits to display after the decimal point.
' @remarks <p>This is used when a specifier of "p" is used with no
' associated digits (ie "p5").</p>
' <p>If the number is less than the precision
' of the value, then the value is rounded to fit within the number of
' digits specified. If omitted the value is rounded to the number of
' digits returned by this property.</p>
' <p>If the specified number of digits is larger than the precision of
' value, then the formatted value will be padded with zeros.</p>
' <p>The default number of digits is 2.</p>
Public Property Get PercentDecimalDigits() As Long
    VerifyLoaded
    PercentDecimalDigits = mProps.PercentDecimalDigits
End Property

''
' Sets the number of digits after the decimal point for percent formatted numbers.
'
' @param RHS The number of digits to display after the decimal point.
' @remarks <p>This is used when a specifier of "p" is used with no
' associated digits (ie "p5").</p>
' <p>If the number is less than the precision
' of the value, then the value is rounded to fit within the number of
' digits specified. If omitted the value is rounded to the number of
' digits returned by this property.</p>
' <p>If the specified number of digits is larger than the precision of
' value, then the formatted value will be padded with zeros.</p>
' <p>The default number of digits is 2.</p>
Public Property Let PercentDecimalDigits(ByVal RHS As Long)
    VerifyWritable
    VerifyLoaded
    If RHS < 0 Or RHS > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99), "PercentDecimalDigits", RHS)
        
    mProps.PercentDecimalDigits = RHS
End Property

''
' Returns the current group sizes array for percent formatted numbers.
'
' @return An array containing grouping sizes for a percent formatted number.
' @remarks A group size is the number of digits that are to be grouped together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Get PercentGroupsSizes() As Variant
    VerifyLoaded
    PercentGroupSizes = mProps.PercentGroupSizes
End Property

''
' Sets the group sizes for percent formatted numbers.
'
' @param RHS A vbLong array containing the grouping sizes.
' @remarks A group size is the number of digits that are to be groupsed together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Let PercentGroupSizes(ByRef RHS As Variant)
    VerifyWritable
    VerifyLoaded
    If cArray.IsNull(RHS) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "PercentGroupSizes")
    
    VerifyGroupSizes RHS
    mProps.PercentGroupSizes = RHS
End Property

''
' Returns the selected negative pattern used in percent number formatting.
'
' @return Then negative number pattern to be used in formating negative numbers.
'
Public Property Get PercentNegativePattern() As PercentNegativePatterns
    VerifyLoaded
    PercentNegativePattern = mProps.PercentNegativePattern
End Property

''
' Sets the negative pattern to be used in formatting negative percent numbers.
'
' @param RHS The pattern to be used.
'
Public Property Let PercentNegativePattern(ByVal RHS As PercentNegativePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.PercentNegativePattern = RHS
End Property

''
' Returns the group separator for a percent formatted number.
'
' @return The group separator symbol.
' @remarks Generally a group separator is a comman ','. However,
' different cultures may use different symbols.
'
Public Property Get PercentGroupSeparator() As String
    VerifyLoaded
    PercentGroupSeparator = mProps.PercentGroupSeparator
End Property

''
' Sets the group separator for a percent formatted number.
'
' @param RHS The symbol used as a group separator defined by the group sizes.
'
Public Property Let PercentGroupSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentGroupSeparator = RHS
End Property

''
' Returns the percent symbol used in percent formatted numbers.
'
' @return The percent symbole, usually '%'.
'
Public Property Get PercentSymbol() As String
    VerifyLoaded
    PercentSymbol = mProps.PercentSymbol
End Property

''
' Set the percent symbol used in percent formatted numbers.
'
' @param RHS The symbol to be used in percent formatted numbers.
'
Public Property Let PercentSymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentSymbol = RHS
End Property

''
' Returns the pattern to be used in formatting positive percent numbers.
'
' @return The pattern used in formatting positive percent numbers.
'
Public Property Get PercentPositivePattern() As PercentPositivePatterns
    VerifyLoaded
    PercentPositivePattern = mProps.PercentPositivePattern
End Property

''
' Sets the pattern to be used in formatting positive percent numbers.
'
' @param RHS The pattern to be used.
'
Public Property Let PercentPositivePattern(ByVal RHS As PercentPositivePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.PercentPositivePattern = RHS
End Property

''
' Returns the decimal separator used in formatting numbers into currency.
'
' @return The decimal separator character.
'
Public Property Get CurrencyDecimalSeparator() As String
    VerifyLoaded
    CurrencyDecimalSeparator = mProps.CurrencyDecimalSeparator
End Property

''
' Sets the decimal separator used in formatting numbers into currency.
'
' @param RHS The decimal separator to be used in formatting.
'
Public Property Let CurrencyDecimalSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyDecimalSeparator = RHS
End Property

''
' Returns the default number of digits after the decimal separator.
'
' @return The default number of digits after the decimal separator.
'
Public Property Get CurrencyDecimalDigits() As Long
    VerifyLoaded
    CurrencyDecimalDigits = mProps.CurrencyDecimalDigits
End Property

''
' Sets the default number of digits to be shown after the decimal separator.
'
' @param RHS The default number of digits to be shown after the decimal separator.
'
Public Property Let CurrencyDecimalDigits(ByVal RHS As Long)
    VerifyWritable
    VerifyLoaded
    If RHS < 0 Or RHS > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99), "CurrencyDecimalDigits", RHS)
    mProps.CurrencyDecimalDigits = RHS
End Property

''
' Returns an array of longs which represent a comma separator group size.
'
' @return An array of longs.
' @remarks <p>Each element represents the number of digits that will be grouped
' together before a comma separator is inserted. Each element is used once
' to specify the group size, then the next element is used for the next group
' of digits, until the last group size. The last group size is used for the
' remainder of digits to be formatted.</p>
' <p>If a 0 is specified for a group size, it can only be the last size in the array.
' A 0 as the last group size will prevent the remaining digits from being separated.</p>
'
Public Property Get CurrencyGroupsSizes() As Variant
    VerifyLoaded
    CurrencyGroupSizes = mProps.CurrencyGroupSizes
End Property

''
' Sets an array of longs which represent a set of group sizes before being comma separated.
'
' @param RHS An array of longs that are the group sizes.
' @remarks <p>Each element represents the number of digits that will be grouped
' together before a comma separator is inserted. Each element is used once
' to specify the group size, then the next element is used for the next group
' of digits, until the last group size. The last group size is used for the
' remainder of digits to be formatted.</p>
' <p>If a 0 is specified for a group size, it can only be the last size in the array.
' A 0 as the last group size will prevent the remaining digits from being separated.</p>
'
Public Property Let CurrencyGroupSizes(ByRef RHS As Variant)
    VerifyWritable
    VerifyLoaded
    If cArray.IsNull(RHS) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "CurrencyGroupSizes")
    
    VerifyGroupSizes RHS
    mProps.CurrencyGroupSizes = RHS
End Property

''
' The numeric pattern to be used when formatting negative values.
'
' @return The negative pattern used.
'
Public Property Get CurrencyNegativePattern() As CurrencyNegativePatterns
    VerifyLoaded
    CurrencyNegativePattern = mProps.CurrencyNegativePattern
End Property

''
' Sets the numeric pattern to be used when formatting negative values.
'
' @param RHS The pattern to be used when formatting.
'
Public Property Let CurrencyNegativePattern(ByVal RHS As CurrencyNegativePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyNegativePattern = RHS
End Property

''
' Returns the string to be use as a group separator.
'
' @return The group separator string.
' @remarks Generally a comma is the group separator. However, in some
' parts of the world, a comma is not used.
'
Public Property Get CurrencyGroupSeparator() As String
    VerifyLoaded
    CurrencyGroupSeparator = mProps.CurrencyGroupSeparator
End Property

''
' Sets the group separator to be used when separating digits.
'
' @param RHS The separator to be used when separating digits.
' @remarks Generally a comma is the group separator. However, in some
' parts of the world, a comma is not used.
'
Public Property Let CurrencyGroupSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyGroupSeparator = RHS
End Property

''
' Returns the currency symbol.
'
' @return The currency symbol.
' @remarks In the US the symbol is the dollar sign '$', in other parts
' of the world, other symbols are used.
'
Public Property Get CurrencySymbol() As String
    VerifyLoaded
    CurrencySymbol = mProps.CurrencySymbol
End Property

''
' Sets the currency symbol to be used.
'
' @param RHS The currency symbol to be used when formatting currency.
' @remarks In the US the symbol is the dollar sign '$', in other parts
' of the world, other symbols are used.
'
Public Property Let CurrencySymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencySymbol = RHS
End Property

''
' Returns the formatting pattern to be used for positive currency values.
'
' @return The formatting pattern for positive currency values.
'
Public Property Get CurrencyPositivePattern() As CurrencyPositivePatterns
    VerifyLoaded
    CurrencyPositivePattern = mProps.CurrencyPositivePattern
End Property

''
' Sets the formatting pattern to be used for positive currency values.
'
' @param RHS The formatting pattern to be used with positive currency values.
'
Public Property Let CurrencyPositivePattern(ByVal RHS As CurrencyPositivePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyPositivePattern = RHS
End Property

''
' Returns the separator character for groups of numbers.
'
' @return The character that is used as a separator between groups of numbers.
' This is typically a comma.
'
Public Property Get NumberGroupSeparator() As String
    VerifyLoaded
    NumberGroupSeparator = mProps.NumberGroupSeparator
End Property

''
' Sets the separator charactor for a group of numbers.
'
' @param RHS The character to be used when separating groups of numbers.
' This is typically a comma.
'
Public Property Let NumberGroupSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NumberGroupSeparator = RHS
    EnsureCapacity
End Property

''
' Returns the pattern to be used when formatting negative numbers.
'
' @return The pattern to be used.
'
Public Property Get NumberNegativePattern() As NumberNegativePatterns
    VerifyLoaded
    NumberNegativePattern = mProps.NumberNegativePattern
End Property

''
' Sets the pattern to be used when formatting negative numbers.
'
' @param RHS The new pattern to be used.
'
Public Property Let NumberNegativePattern(ByVal RHS As NumberNegativePatterns)
    VerifyWritable
    VerifyLoaded
    If RHS < 0 Or RHS > 4 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "NumberNegativePattern", RHS)
    mProps.NumberNegativePattern = RHS
End Property

''
' Returns the current group sizes array for formatted numbers.
'
' @return An array containing grouping sizes for a formatted number.
' @remarks A group size is the number of digits that are to be grouped together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Get NumberGroupSizes() As Variant
    VerifyLoaded
    NumberGroupSizes = mProps.NumberGroupSizes
End Property

''
' Sets the group sizes for formatted numbers.
'
' @param RHS A vbLong array containing the grouping sizes.
' @remarks A group size is the number of digits that are to be grouped together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Let NumberGroupSizes(ByRef RHS As Variant)
    VerifyWritable
    VerifyLoaded
    If cArray.IsNull(RHS) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "NumberGroupSizes")
    
    VerifyGroupSizes RHS
    mProps.NumberGroupSizes = RHS
End Property

''
' The default number of digits to be appended after the decimal point if one is used.
'
' @return The number of digits to be appended after the decimal point if one is used.
'
Public Property Get NumberDecimalDigits() As Long
    VerifyLoaded
    NumberDecimalDigits = mProps.NumberDecimalDigits
End Property

''
' Sets the default number of digits to be appended after the decimal point.
'
' @param RHS The number of digits to be appended after the decimal point.
'
Public Property Let NumberDecimalDigits(ByVal RHS As Long)
    VerifyWritable
    VerifyLoaded
    If RHS < 0 Or RHS > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99), "NumberDecimalDigits", RHS)
    mProps.NumberDecimalDigits = RHS
End Property

''
' Returns the character used as a decimal point.
'
' @return The decimal point character.
'
Public Property Get NumberDecimalSeparator() As String
    VerifyLoaded
    NumberDecimalSeparator = mProps.NumberDecimalSeparator
End Property

''
' Sets the character to be used as the decimal point.
'
' @param RHS The new character to be used as the decimal point.
'
Public Property Let NumberDecimalSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NumberDecimalSeparator = RHS
    EnsureCapacity
End Property

''
' Returns the negative symbol for negative numbers.
'
' @return The current negative symbol.
'
Public Property Get NegativeSign() As String
    VerifyLoaded
    NegativeSign = mProps.NegativeSign
End Property

''
' Sets the negative symbol to be used for negative numbers.
'
' @param RHS The new negative symbol.
'
Public Property Let NegativeSign(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NegativeSign = RHS
    EnsureCapacity
End Property

''
' Returns the current positive symbol.
'
' @return The current positive symbol.
'
Public Property Get PositiveSign() As String
    VerifyLoaded
    PositiveSign = mProps.PositiveSign
End Property

''
' Sets the positive symbol.
'
' @param RHS The new positive symbol.
'
Public Property Let PositiveSign(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PositiveSign = RHS
End Property

''
' Format a numeric value with the proper formatting applied to the
' value type. Bytes, Integer, and Longs are formatted as Longs. singles and
' Doubles are formatted individually.
'
' @param value The numeric value to be formatte to the string representation.
' @param fmt The formatting to perform on the value.
' @return The formatted value as a string.
'
Public Function Format(ByRef Value As Variant, Optional ByRef fmt As String) As String
    VerifyLoaded
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            Format = FormatInt32(Value, fmt)
        Case vbDouble
            Format = FormatDouble(Value, fmt)
        Case vbSingle
            Format = FormatSingle(Value, fmt)
        Case Else
            Format = Value
    End Select
End Function

''
' Returns a format provider of the requested type.
'
' @param FormatType The name of the format provider.
' @return The requested format provider, or Nothing if no format provider was found.
'
Public Function GetFormat(ByVal FormatType As String) As Object
    If LCase$(FormatType) = "numberformatinfo" Then Set GetFormat = Me
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub CloneHelper(ByRef props As PropsType)
    mProps = props
End Sub

Friend Sub Load(ByVal LCID As Long, ByVal UseUserOverride As Boolean)
    If Not CultureTable.IsLoaded Then _
        Throw Cor.NewNotSupportedException("Cannot load NumberFormatInfo without culture table.")
    
    LoadCommon LCID
    
    If UseUserOverride Then
        LoadUserOverride LCID
    Else
        LoadFromCultureTable LCID
    End If
    EnsureCapacity
    mLoaded = True
End Sub

Private Sub VerifyLoaded()
    If mLoaded Then Exit Sub
    Load INVARIANT_LCID, False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUserOverride(ByVal LCID As Long)
    With mProps
        .NumberDecimalSeparator = GetLocaleString(LCID, LOCALE_SDECIMAL)
        .NumberDecimalDigits = GetLocaleLong(LCID, LOCALE_IDIGITS)
        .NumberGroupSeparator = GetLocaleString(LCID, LOCALE_STHOUSAND)
        .NumberGroupSizes = ParseGroupSizes(GetLocaleString(LCID, LOCALE_SGROUPING))
        .NumberNegativePattern = GetLocaleLong(LCID, LOCALE_INEGNUMBER)
        .PercentDecimalSeparator = .NumberDecimalSeparator
        .PercentDecimalDigits = .NumberDecimalDigits
        .PercentGroupSeparator = .NumberGroupSeparator
        .PercentGroupSizes = .NumberGroupSizes
        .PercentNegativePattern = .NumberNegativePattern
        .CurrencyDecimalSeparator = GetLocaleString(LCID, LOCALE_SMONDECIMALSEP)
        .CurrencyDecimalDigits = GetLocaleLong(LCID, LOCALE_ICURRDIGITS)
        .CurrencyGroupSeparator = GetLocaleString(LCID, LOCALE_SMONTHOUSANDSEP)
        .CurrencyGroupSizes = ParseGroupSizes(GetLocaleString(LCID, LOCALE_SMONGROUPING))
        .CurrencyNegativePattern = GetLocaleLong(LCID, LOCALE_INEGCURR)
        .CurrencyPositivePattern = GetLocaleLong(LCID, LOCALE_ICURRENCY)
        .CurrencySymbol = GetLocaleString(LCID, LOCALE_SCURRENCY)
        .NegativeSign = GetLocaleString(LCID, LOCALE_SNEGATIVESIGN)
    End With
End Sub

Private Sub LoadFromCultureTable(ByVal LCID As Long)
    With mProps
        .NumberDecimalDigits = CultureTable.GetNumber(LCID, INUMBERDECIMALDIGITS)
        .NumberDecimalSeparator = CultureTable.GetString(LCID, SNUMBERDECIMALSEPARATOR)
        .NumberGroupSeparator = CultureTable.GetString(LCID, SNUMBERGROUPSEPARATOR)
        .NumberGroupSizes = CultureTable.GetNumberArray(LCID, SNUMBERGROUPSIZES)
        .NumberNegativePattern = CultureTable.GetNumber(LCID, INUMBERNEGATIVEPATTERN)
        .PercentDecimalDigits = CultureTable.GetNumber(LCID, IPERCENTDECIMALDIGITS)
        .PercentDecimalSeparator = CultureTable.GetString(LCID, SPERCENTDECIMALSEPARATOR)
        .PercentGroupSeparator = CultureTable.GetString(LCID, SPERCENTGROUPSEPARATOR)
        .PercentGroupSizes = CultureTable.GetNumberArray(LCID, SPERCENTGROUPSIZES)
        .PercentNegativePattern = CultureTable.GetNumber(LCID, IPERCENTNEGATIVEPATTERN)
        .CurrencyDecimalDigits = CultureTable.GetNumber(LCID, ICURRENCYDECIMALDIGITS)
        .CurrencyDecimalSeparator = CultureTable.GetString(LCID, SCURRENCYDECIMALSEPARATOR)
        .CurrencyGroupSeparator = CultureTable.GetString(LCID, SCURRENCYGROUPSEPARATOR)
        .CurrencyGroupSizes = CultureTable.GetNumberArray(LCID, SCURRENCYGROUPSIZES)
        .CurrencyNegativePattern = CultureTable.GetNumber(LCID, ICURRENCYNEGATIVEPATTERN)
        .CurrencyPositivePattern = CultureTable.GetNumber(LCID, ICURRENCYPOSITIVEPATTERN)
        .CurrencySymbol = CultureTable.GetString(LCID, SCURRENCYSYMBOL)
        .NegativeSign = CultureTable.GetString(LCID, SNEGATIVESIGN)
    End With
End Sub

Private Sub LoadCommon(ByVal LCID As Long)
    With mProps
        .NaNSymbol = CultureTable.GetString(LCID, SNANSYMBOL)
        .PositiveInfinitySymbol = CultureTable.GetString(LCID, SPOSITIVEINFINITYSYMBOL)
        .NegativeInfinitySymbol = CultureTable.GetString(LCID, SNEGATIVEINFINITYSYMBOL)
        .PercentPositivePattern = CultureTable.GetNumber(LCID, IPERCENTPOSITIVEPATTERN)
        .PercentSymbol = CultureTable.GetString(LCID, SPERCENTSYMBOL)
        .PositiveSign = CultureTable.GetString(LCID, SPOSITIVESIGN)
    End With
End Sub

Private Function ParseGroupSizes(ByRef src As String) As Long()
    Dim i As Long
    Dim groups() As String
    Dim ret() As Long
    
    groups = Split(src, ";")
    ReDim ret(0 To UBound(groups) - 1)
    
    For i = 0 To UBound(groups) - 1 ' we don't use the last element.
        ret(i) = CLng(groups(i))
    Next i
    ParseGroupSizes = ret
End Function

Private Function FormatSingle(ByVal Value As Single, ByRef Format As String) As String
    Dim Precision As Long
    Dim fmt As Long
    
    fmt = ParseFormatInfo(Format, Precision)
    Select Case fmt And &HFFDF
        Case GENERAL_FORMAT
            If Precision < 0 Then Precision = SINGLE_PRECISION
            DoubleToNumber Value, Precision
            If Not IsValidNumber(FormatSingle) Then Exit Function
            If Precision > 7 Then Precision = 9
            RoundNumber Precision
            FormatGeneralDouble mNumber.Precision, fmt - 2
        
        Case FIXED_FORMAT
            If Precision < 0 Then Precision = 2
            DoubleToNumber Value, SINGLE_PRECISION
            If Not IsValidNumber(FormatSingle) Then Exit Function
            RoundNumber Precision + mNumber.Scale
            FormatFixedDouble Precision
            
        Case EXPONENT_FORMAT
            If Precision < 0 Then Precision = 6
            DoubleToNumber Value, SINGLE_PRECISION
            If Not IsValidNumber(FormatSingle) Then Exit Function
            RoundNumber Precision + 1
            FormatScientific Precision, fmt
            
        Case DECIMAL_FORMAT
            Throw Cor.NewFormatException("Invalid format specifier.")
       
        Case NUMBER_FORMAT
            DoubleToNumber Value, SINGLE_PRECISION
            If Not IsValidNumber(FormatSingle) Then Exit Function
            If Precision < 0 Then Precision = mProps.NumberDecimalDigits
            Dim i As Long
            i = mNumber.Scale + Precision
            If i > SINGLE_PRECISION Then i = SINGLE_PRECISION + 2
            RoundNumber i
            FormatNumber Precision
            
        Case CURRENCY_FORMAT
            DoubleToNumber Value, SINGLE_PRECISION
            If Not IsValidNumber(FormatSingle) Then Exit Function
            If Precision < 0 Then Precision = mProps.NumberDecimalDigits
            i = mNumber.Scale + Precision
            If i > SINGLE_PRECISION Then i = SINGLE_PRECISION + 2
            RoundNumber i
            FormatCurrency Precision

        Case PERCENT_FORMAT
            DoubleToNumber Value, SINGLE_PRECISION
            If Not IsValidNumber(FormatSingle) Then Exit Function
            If Precision < 0 Then Precision = mProps.PercentDecimalDigits
            mNumber.[Scale] = mNumber.Scale + 2
            RoundNumber Precision + mNumber.Scale
            FormatPercent Precision

        Case Else
            FormatSingle = VBA.Format$(Value, Format)
            Exit Function
    End Select
    If mBufferPos = 0 Then
        FormatSingle = "0"
    Else
        FormatSingle = SysAllocStringLen(mBuffer.SA.pvData, mBufferPos)
    End If
End Function

Private Function FormatDouble(ByVal Value As Double, ByRef Format As String) As String
    Dim Precision As Long
    Dim fmt As Long
    
    fmt = ParseFormatInfo(Format, Precision)
    Select Case fmt And &HFFDF
        Case GENERAL_FORMAT
            If Precision < 0 Then Precision = DOUBLE_PRECISION
            DoubleToNumber Value, Precision
            If Not IsValidNumber(FormatDouble) Then Exit Function
            RoundNumber Precision
            FormatGeneralDouble Precision, fmt - 2

        Case FIXED_FORMAT
            If Precision < 0 Then Precision = 2
            DoubleToNumber Value, DOUBLE_PRECISION
            If Not IsValidNumber(FormatDouble) Then Exit Function
            RoundNumber Precision + mNumber.Scale
            FormatFixedDouble Precision
            
        Case EXPONENT_FORMAT
            If Precision < 0 Then Precision = 6
            DoubleToNumber Value, DOUBLE_PRECISION
            If Not IsValidNumber(FormatDouble) Then Exit Function
            RoundNumber Precision + 1
            FormatScientific Precision, fmt
            
        Case DECIMAL_FORMAT
            Throw Cor.NewFormatException("Invalid format specifier.")
            
        Case NUMBER_FORMAT
            DoubleToNumber Value, DOUBLE_PRECISION
            If Not IsValidNumber(FormatDouble) Then Exit Function
            If Precision < 0 Then Precision = mProps.NumberDecimalDigits
            RoundNumber Precision + mNumber.Scale
            FormatNumber Precision
            
        Case CURRENCY_FORMAT
            DoubleToNumber Value, DOUBLE_PRECISION
            If Not IsValidNumber(FormatDouble) Then Exit Function
            If Precision < 0 Then Precision = mProps.CurrencyDecimalDigits
            RoundNumber Precision + mNumber.Scale
            FormatCurrency Precision
            
        Case PERCENT_FORMAT
            DoubleToNumber Value, DOUBLE_PRECISION
            If Not IsValidNumber(FormatDouble) Then Exit Function
            If Precision < 0 Then Precision = mProps.PercentDecimalDigits
            mNumber.[Scale] = mNumber.Scale + 2
            RoundNumber Precision + mNumber.Scale
            FormatPercent Precision
            
        Case Else
            FormatDouble = VBA.Format$(Value, Format)
            Exit Function
    End Select
    If mBufferPos = 0 Then
        FormatDouble = "0"
    Else
        FormatDouble = SysAllocStringLen(mBuffer.SA.pvData, mBufferPos)
    End If
End Function

Private Function FormatInt32(ByRef Value As Variant, ByRef Format As String) As String
    Dim digits As Long
    Dim fmt As Long
    
    fmt = ParseFormatInfo(Format, digits)
    Select Case fmt And &HFFDF
        Case GENERAL_FORMAT
            Int32ToNumber Value
            If digits < 0 Then digits = INTEGER_PRECISION
            FormatGeneralInt32 digits, fmt - 2
            
        Case FIXED_FORMAT
            Int32ToNumber Value
            If digits < 0 Then digits = mProps.NumberDecimalDigits
            FormatFixedInt32 digits
            
        Case EXPONENT_FORMAT
            Int32ToNumber Value
            If digits < 0 Then digits = 6
            RoundNumber digits + 1
            FormatScientific digits, fmt
            
        Case DECIMAL_FORMAT
            Int32ToNumber Value
            FormatDecimal digits
            
        Case HEX_FORMAT
            If VarType(Value) = vbInteger Then
                ' to allow integers to be formatted to 4 characters
                ' instead of 8 like longs can be.
                FormatHex CLng(Value) And &HFFFF&, digits, fmt - 33
            Else
                FormatHex Value, digits, fmt - 33
            End If
            FormatInt32 = SysAllocStringLen(mBuffer.SA.pvData + mBufferPos * 2, mBufferCapacity - mBufferPos)
            Exit Function
        
        Case NUMBER_FORMAT
            Int32ToNumber Value
            If digits < 0 Then digits = mProps.NumberDecimalDigits
            FormatNumber digits
        
        Case CURRENCY_FORMAT
            Int32ToNumber Value
            If digits < 0 Then digits = mProps.CurrencyDecimalDigits
            FormatCurrency digits
            
        Case PERCENT_FORMAT
            Int32ToNumber Value
            If mNumber.Scale > 0 Then mNumber.[Scale] = mNumber.Scale + 2
            If digits < 0 Then digits = mProps.PercentDecimalDigits
            FormatPercent digits
            
        Case Else
            FormatInt32 = VBA.Format$(Value, Format)
            Exit Function
    End Select
    If mBufferPos = 0 Then
        FormatInt32 = "0"
    Else
        FormatInt32 = SysAllocStringLen(mBuffer.SA.pvData, mBufferPos)
    End If
End Function

Private Sub FormatNumber(ByVal decimalplaces As Long)
    Dim Pattern As String
    Dim i As Long
    Dim ch As Integer
    
    If mNumber.IsNegative Then
        Pattern = mNumberNegativePatterns(mProps.NumberNegativePattern)
    Else
        Pattern = "#"
    End If
    mFormatInfo.SA.pvData = StrPtr(Pattern)
    For i = 0 To Len(Pattern) - 1
        ch = mFormatInfo.Data(i)
        Select Case ch
            Case CHAR_POUND
                With mProps
                    FormatGroups decimalplaces, .NumberGroupSizes, .NumberGroupSeparator, .NumberDecimalSeparator
                End With
            Case CHAR_MINUS:    AppendString mProps.NegativeSign
            Case Else
                mBuffer.Data(mBufferPos) = ch
                mBufferPos = mBufferPos + 1
        End Select
    Next i
End Sub

Private Sub FormatCurrency(ByVal decimalplaces As Long)
    Dim Pattern As String
    Dim i As Long
    Dim ch As Integer
    
    If mNumber.IsNegative Then
        Pattern = mCurrencyNegativePatterns(mProps.CurrencyNegativePattern)
    Else
        Pattern = mCurrencyPositivePatterns(mProps.CurrencyPositivePattern)
    End If
    mFormatInfo.SA.pvData = StrPtr(Pattern)
    For i = 0 To Len(Pattern) - 1
        ch = mFormatInfo.Data(i)
        Select Case ch
            Case CHAR_POUND
                With mProps
                    FormatGroups decimalplaces, .CurrencyGroupSizes, .CurrencyGroupSeparator, .CurrencyDecimalSeparator
                End With
            Case CHAR_MINUS:    AppendString mProps.NegativeSign
            Case CHAR_DOLLAR:   AppendString mProps.CurrencySymbol
            Case Else
                mBuffer.Data(mBufferPos) = ch
                mBufferPos = mBufferPos + 1
        End Select
    Next i
End Sub

Private Sub FormatPercent(ByVal decimalplaces As Long)
    Dim Pattern As String
    Dim i As Long
    Dim ch As Integer
    
    If mNumber.IsNegative Then
        Pattern = mPercentNegativePatterns(mProps.PercentNegativePattern)
    Else
        Pattern = mPercentPositivePatterns(mProps.PercentPositivePattern)
    End If
    mFormatInfo.SA.pvData = StrPtr(Pattern)
    For i = 0 To Len(Pattern) - 1
        ch = mFormatInfo.Data(i)
        Select Case ch
            Case CHAR_POUND
                With mProps
                    FormatGroups decimalplaces, .PercentGroupSizes, .PercentGroupSeparator, .PercentDecimalSeparator
                End With
            Case CHAR_MINUS:     AppendString mProps.NegativeSign
            Case CHAR_PERCENT:   AppendString mProps.PercentSymbol
            Case Else
                mBuffer.Data(mBufferPos) = ch
                mBufferPos = mBufferPos + 1
        End Select
    Next i
End Sub

Private Sub FormatGroups(ByVal decimalplaces As Long, ByRef GroupSizes() As Long, ByRef GroupSeparator As String, ByRef DecimalSeparator As String)
    Dim i As Long
    Dim j As Long
    Dim ch As Integer
    Dim remainingdigits As Long
    Dim digitcount As Long
    Dim Size As Long
    
    With mNumber
        If .Scale > 0 Then
            For i = LBound(GroupSizes) To UBound(GroupSizes) - 1
                j = j + GroupSizes(i)
            Next i
            
            remainingdigits = .Scale - j
            
            If UBound(GroupSizes) >= 0 Then Size = GroupSizes(UBound(GroupSizes))
            If remainingdigits = Size Then
                digitcount = -1
            ElseIf Size > 0 Then
                digitcount = remainingdigits Mod Size
                If digitcount = 0 Then digitcount = Size
            Else
                digitcount = remainingdigits
            End If
            i = 0
            ch = .digits(i)
            Do While remainingdigits > 0
                If digitcount = 0 Then
                    AppendString GroupSeparator
                    digitcount = Size
                End If
                If ch = 0 Then
                    mBuffer.Data(mBufferPos) = CHAR_0
                Else
                    mBuffer.Data(mBufferPos) = ch
                    ch = .digits(i + 1)
                End If
                i = i + 1
                mBufferPos = mBufferPos + 1
                digitcount = digitcount - 1
                remainingdigits = remainingdigits - 1
            Loop
        Else
            mBuffer.Data(mBufferPos) = CHAR_0
            mBufferPos = mBufferPos + 1
        End If
        
        If (i = 0 Or ch <> 0) Or (i < .Scale And ch = 0) Then
            Size = UBound(GroupSizes)
            digitcount = 0
            ch = .digits(i)
            For j = i To .Scale - 1
                If digitcount = 0 Then
                    AppendString GroupSeparator
                    Size = Size - 1
                    If Size < 0 Then Size = 0
                    digitcount = GroupSizes(Size)
                End If
                If ch = 0 Then
                    mBuffer.Data(mBufferPos) = CHAR_0
                Else
                    mBuffer.Data(mBufferPos) = ch
                    ch = .digits(i + 1)
                End If
                i = i + 1
                mBufferPos = mBufferPos + 1
                digitcount = digitcount - 1
            Next j
        End If
        
        If decimalplaces > 0 Then
            AppendString DecimalSeparator
            i = .Scale
            Do While i < 0
                mBuffer.Data(mBufferPos) = CHAR_0
                mBufferPos = mBufferPos + 1
                i = i + 1
                decimalplaces = decimalplaces - 1
            Loop
            If ch <> 0 Then ch = .digits(i)
            Do While decimalplaces > 0
                If ch = 0 Then
                    mBuffer.Data(mBufferPos) = CHAR_0
                Else
                    mBuffer.Data(mBufferPos) = ch
                    i = i + 1
                    ch = .digits(i)
                End If
                mBufferPos = mBufferPos + 1
                decimalplaces = decimalplaces - 1
            Loop
        End If
    End With
End Sub

Private Sub FormatHex(ByVal Value As Long, ByVal digits As Long, ByVal caseoffset As Long)
    Dim ch As Integer
    
    mBufferPos = mBufferCapacity - 1
    Do While Value <> 0
        ch = Value And &HF
        If ch > 9 Then
            mBuffer.Data(mBufferPos) = ch + caseoffset
        Else
            mBuffer.Data(mBufferPos) = CHAR_0 + ch
        End If
        mBufferPos = mBufferPos - 1
        Value = Helper.ShiftRight(Value, 4)
    Loop
    Do While mBufferPos >= mBufferCapacity - digits
        mBuffer.Data(mBufferPos) = CHAR_0
        mBufferPos = mBufferPos - 1
    Loop
    mBufferPos = mBufferPos + 1
End Sub

Private Sub FormatDecimal(ByVal digits As Long)
    Dim i As Long
    
    With mNumber
        If .IsNegative Then AppendString mProps.NegativeSign
        Do While digits > .Scale
            mBuffer.Data(mBufferPos) = CHAR_0
            mBufferPos = mBufferPos + 1
            digits = digits - 1
        Loop
        For i = 0 To .Scale - 1
            mBuffer.Data(mBufferPos) = .digits(i)
            mBufferPos = mBufferPos + 1
        Next i
    End With
End Sub

Private Sub FormatScientific(ByVal digits As Long, ByVal expchar As Integer)
    Dim i As Long
    Dim ch As Integer
    
    With mNumber
        If .IsNegative Then AppendString mProps.NegativeSign
        Do While digits >= 0
            If i = 1 Then AppendString mProps.NumberDecimalSeparator
            ch = .digits(i)
            If ch = 0 Then
                mBuffer.Data(mBufferPos) = CHAR_0
            Else
                mBuffer.Data(mBufferPos) = ch
                i = i + 1
            End If
            mBufferPos = mBufferPos + 1
            digits = digits - 1
        Loop
        FormatExponent .Scale - 1, expchar, 3
    End With
End Sub

Private Sub RoundNumber(ByVal digits As Long)
    If digits <= 0 Then Exit Sub
    With mNumber
        If .digits(digits) >= CHAR_5 Then
            Do While digits > 0
                If .digits(digits - 1) <> CHAR_9 Then Exit Do
                digits = digits - 1
            Loop
            If digits > 0 Then
                .digits(digits - 1) = .digits(digits - 1) + 1
            Else
                .Scale = .Scale + 1
                .digits(digits) = CHAR_1
                digits = 1
            End If
        Else
            Do While digits > 0
                If .digits(digits - 1) <> CHAR_0 Then Exit Do
                digits = digits - 1
            Loop
        End If
        If digits = 0 Then
            .Scale = 0
            .IsNegative = BOOL_FALSE
        End If
        .digits(digits) = vbNullPtr
    End With
End Sub

Private Function ParseFormatInfo(ByRef fmt As String, ByRef digits As Long) As Long
    If LenB(fmt) > 0 Then
        With mFormatInfo
            .SA.pvData = StrPtr(fmt)
            
            Dim ch As Integer
            Dim i As Long
            Dim j As Long
            Dim dig As Integer
            ch = .Data(0)
            Select Case ch
                Case CHAR_LOWER_A To CHAR_LOWER_Z, CHAR_UPPER_A To CHAR_UPPER_Z
                    For i = 1 To Len(fmt) - 1
                        dig = .Data(i)
                        Select Case dig
                            Case CHAR_0 To CHAR_9
                                j = j * 10 + dig - CHAR_0
                            Case Else
                                Exit Function
                        End Select
                    Next i
                    If i = 1 Then j = -1
                    digits = j
                    ParseFormatInfo = ch
                    Exit Function
                Case Else
                    Exit Function
            End Select
        End With
    End If
    digits = -1
    ParseFormatInfo = GENERAL_FORMAT
End Function

Private Sub Int32ToNumber(ByVal Value As Long)
    mBufferPos = 0
    With mNumber
        .DigitsSA.pvData = mPtrNumberBuffer
        If Value < 0 Then
            .IsNegative = BOOL_TRUE
            Value = -Value
        Else
            .IsNegative = BOOL_FALSE
        End If
        
        Dim i As Long
        Do While Value > 0
            .digits(i) = (Value Mod 10) + CHAR_0
            Value = Value \ 10
            i = i + 1
        Loop
        .Precision = i
        .Scale = i
        .digits(i) = vbNullPtr
        Dim j As Long
        Dim t As Long
        Do While j < i
            i = i - 1
            t = .digits(j)
            .digits(j) = .digits(i)
            .digits(i) = t
            j = j + 1
        Loop
    End With
End Sub

Private Sub DoubleToNumber(ByVal Value As Double, ByVal Precision As Long)
    mBufferPos = 0
    mNumber.Precision = Precision
    With AsDLong(Value)
        If (.HiDWord And &H7FF00000) = &H7FF00000 Then
            If (.HiDWord And &HFFFFF) Or (.LoDWord) Then
                mNumber.[Scale] = SCALE_NAN
            Else
                mNumber.[Scale] = SCALE_INF
                mNumber.IsNegative = (Value < 0#)
            End If
        Else
            mNumber.DigitsSA.pvData = Helper.Ecvt(Value, Precision, mNumber.Scale, mNumber.IsNegative)
        End If
    End With
End Sub

Private Sub FormatGeneralInt32(ByVal digits As Long, ByVal expchar As Integer)
    Dim hasexponent As Boolean
    
    With mNumber
        If .IsNegative Then AppendString mProps.NegativeSign
        If digits < .Precision Then
            RoundNumber digits
            .Scale = 1
            hasexponent = True
        Else
            digits = .Precision
        End If
        
        Dim i As Long
        For i = 0 To digits - 1
            If i = .Scale Then AppendString mProps.NumberDecimalSeparator
            mBuffer.Data(mBufferPos) = .digits(i)
            mBufferPos = mBufferPos + 1
        Next i
        If hasexponent Then FormatExponent .Precision - 1, expchar, 2
    End With
End Sub

Private Sub FormatGeneralDouble(ByVal digits As Long, ByVal expchar As Integer)
    Dim hasexponent As Boolean
    Dim dp As Long
    Dim i As Long
    Dim ch As Integer
    
    With mNumber
        If .IsNegative Then AppendString mProps.NegativeSign
        dp = .Scale
        If digits < dp Or dp < -3 Then
            hasexponent = True
            dp = 1
        ElseIf dp < 0 Then
            mBuffer.Data(mBufferPos) = CHAR_0
            mBufferPos = mBufferPos + 1
            dp = 0
        End If
        
        ch = .digits(i)
        Do While ch <> 0
            If i = dp Then
                AppendString mProps.NumberDecimalSeparator
                If Not hasexponent Then
                    For dp = dp To .Scale + 1 Step -1
                        mBuffer.Data(mBufferPos) = CHAR_0
                        mBufferPos = mBufferPos + 1
                    Next dp
                End If
            End If
            mBuffer.Data(mBufferPos) = ch
            mBufferPos = mBufferPos + 1
            i = i + 1
            ch = .digits(i)
        Loop
        If hasexponent Then
            FormatExponent .Scale - 1, expchar, 2
        Else
            For i = i To .Scale - 1
                mBuffer.Data(mBufferPos) = CHAR_0
                mBufferPos = mBufferPos + 1
            Next i
        End If
    End With
End Sub

Private Sub FormatFixedInt32(ByVal Precision As Long)
    Dim i As Long
    With mNumber
        If .IsNegative Then AppendString mProps.NegativeSign
        For i = 0 To .Scale - 1
            mBuffer.Data(mBufferPos) = .digits(i)
            mBufferPos = mBufferPos + 1
        Next i
        If Precision > 0 Then
            AppendString mProps.NumberDecimalSeparator
            For i = 1 To Precision
                mBuffer.Data(mBufferPos) = CHAR_0
                mBufferPos = mBufferPos + 1
            Next i
        End If
    End With
End Sub

Private Sub FormatFixedDouble(ByVal Precision As Long)
    Dim i As Long
    Dim ch As Integer
    With mNumber
        If .IsNegative Then AppendString mProps.NegativeSign
        If .Scale > 0 Then
            For i = 0 To .Scale - 1
                mBuffer.Data(mBufferPos) = .digits(i)
                mBufferPos = mBufferPos + 1
            Next i
        Else
            mBuffer.Data(mBufferPos) = CHAR_0
            mBufferPos = mBufferPos + 1
            i = .Scale
        End If
        If Precision > 0 Then
            AppendString mProps.NumberDecimalSeparator
            Do While Precision > 0
                If i >= 0 Then ch = .digits(i)
                If ch = 0 Then
                    mBuffer.Data(mBufferPos) = CHAR_0
                    If i < 0 Then i = i + 1
                Else
                    mBuffer.Data(mBufferPos) = ch
                    i = i + 1
                End If
                mBufferPos = mBufferPos + 1
                Precision = Precision - 1
            Loop
        End If
    End With
End Sub

Private Sub AppendString(ByRef stringToAppend As String)
    Dim i As Long
    With mAppendStringBuffer
        .SA.pvData = StrPtr(stringToAppend)
        For i = 0 To Len(stringToAppend) - 1
            mBuffer.Data(mBufferPos) = .Data(i)
            mBufferPos = mBufferPos + 1
        Next i
    End With
End Sub

Private Function IsValidNumber(ByRef retval As String) As Boolean
    Select Case mNumber.Scale
        Case SCALE_NAN
            retval = mProps.NaNSymbol
        Case SCALE_INF
            If mNumber.IsNegative Then
                retval = mProps.NegativeInfinitySymbol
            Else
                retval = mProps.PositiveInfinitySymbol
            End If
        Case Else
            IsValidNumber = True
    End Select
End Function

Private Sub FormatExponent(ByVal places As Long, ByVal expchar As Integer, ByVal mindigits As Long)
    Dim digits As String
    Dim i As Long
    
    mBuffer.Data(mBufferPos) = expchar
    mBufferPos = mBufferPos + 1
    If places < 0 Then
        places = -places
        AppendString mProps.NegativeSign
    Else
        AppendString mProps.PositiveSign
    End If
    digits = places
    i = mindigits - Len(digits)
    Do While i > 0
        mBuffer.Data(mBufferPos) = CHAR_0
        mBufferPos = mBufferPos + 1
        i = i - 1
    Loop
    AppendString digits
End Sub

Private Sub VerifyWritable()
    If mProps.IsReadOnly Then
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_ReadOnly))
    End If
End Sub

Private Sub EnsureCapacity()
    Dim NewCapacity As Long
    NewCapacity = 99 * Len(mProps.NumberGroupSeparator) + 99 + Len(mProps.NegativeSign) + Len(mProps.NumberDecimalSeparator)
    If NewCapacity > mBufferCapacity Then
        CoTaskMemFree mBuffer.SA.pvData
        mBuffer.SA.pvData = CoTaskMemAlloc(NewCapacity * 2)
        If mBuffer.SA.pvData = 0 Then Throw New OutOfMemoryException
        mBufferCapacity = NewCapacity
    End If
End Sub

Private Sub VerifyGroupSizes(ByRef arr As Variant)
    Dim i As Long
    Dim ub As Long
    
    If VarType(arr) <> (vbLong Or vbArray) Then _
        Throw Cor.NewArgumentException("Group sizes array must by of type Long", "Array")
        
    ub = UBound(arr)
    For i = 0 To ub
        Select Case arr(i)
            Case Is < 0, Is > 9:    Throw Cor.NewArgumentException("Group sizes must be between 0 and 9.")
            Case 0:                 If i < ub Then Throw Cor.NewArgumentException("Only last group size can be 0.")
        End Select
    Next i
End Sub

Private Sub ReadLongArray(ByRef a() As Long, ByVal Name As String, ByVal bag As PropertyBag)
    Dim i As Long
    Dim Count As Long
    
    With bag
        Count = .ReadProperty(Name & "_Count", 0)
        If Count = 0 Then
            a = cArray.CreateInstance(ciLong)
            Exit Sub
        End If
        ReDim a(0 To Count - 1)
        For i = 0 To Count - 1
            a(i) = .ReadProperty(Name & "_" & i)
        Next i
    End With
End Sub

Private Sub WriteLongArray(ByRef a() As Long, ByVal Name As String, ByVal bag As PropertyBag)
    Dim i As Long
    
    With bag
        .WriteProperty Name & "_Count", UBound(a) + 1
        For i = 0 To UBound(a)
            .WriteProperty Name & "_" & i, a(i)
        Next i
    End With
End Sub

Private Sub InitPatterns()
    mNumberNegativePatterns = cArray.NewArray(ciString, "(#)", "-#", "- #", "#-", "# -")
    mCurrencyNegativePatterns = cArray.NewArray(ciString, "($#)", "-$#", "$-#", "$#-", "(#$)", "-#$", "#-$", "#$-", "-# $", "-$ #", "# $-", "$ #-", "$ -#", "#- $", "($ #)", "(# $)")
    mCurrencyPositivePatterns = cArray.NewArray(ciString, "$#", "#$", "$ #", "# $")
    mPercentNegativePatterns = cArray.NewArray(ciString, "-# %", "-#%", "-%#")
    mPercentPositivePatterns = cArray.NewArray(ciString, "# %", "#%", "%#")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mProps.PerMilleSymbol = ChrW$(&H2030)
    InitPatterns
    
    With mNumber.DigitsSA
        .cbElements = 1
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    mPtrNumberBuffer = CoTaskMemAlloc(40)
    If mPtrNumberBuffer = 0 Then Throw New OutOfMemoryException
    SAPtr(mNumber.digits) = VarPtr(mNumber.DigitsSA)
    
    InitWordBuffer mBuffer, 0, &H7FFFFFFF
    InitWordBuffer mFormatInfo, 0, &H7FFFFFFF
    InitWordBuffer mAppendStringBuffer, 0, &H7FFFFFFF
End Sub

Private Sub Class_InitProperties()
    'InitDefaults
    'EnsureCapacity
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With mProps
        .CurrencyDecimalDigits = PropBag.ReadProperty(PROP_CURRENCYDECIMALDIGITS)
        .CurrencyDecimalSeparator = PropBag.ReadProperty(PROP_CURRENCYDECIMALSEPARATOR)
        .CurrencyGroupSeparator = PropBag.ReadProperty(PROP_CURRENcYGROUPSEPARATOR)
        ReadLongArray .CurrencyGroupSizes, PROP_CURRENCYGROUPSIZES, PropBag
        .CurrencyNegativePattern = PropBag.ReadProperty(PROP_CURRENCYNEGATIVEPATTERN)
        .CurrencyPositivePattern = PropBag.ReadProperty(PROP_CURRENCYPOSITIVEPATTERN)
        .CurrencySymbol = PropBag.ReadProperty(PROP_CURRENCYSYMBOL)
        .NaNSymbol = PropBag.ReadProperty(PROP_NANSYMBOL)
        .NegativeInfinitySymbol = PropBag.ReadProperty(PROP_NEGATIVEINFINITYSYMBOL)
        .NegativeSign = PropBag.ReadProperty(PROP_NEGATIVESIGN)
        .NumberDecimalDigits = PropBag.ReadProperty(PROP_NUMBERDECIMALDIGITS)
        .NumberDecimalSeparator = PropBag.ReadProperty(PROP_NUMBERDECIMALSEPARATOR)
        .NumberGroupSeparator = PropBag.ReadProperty(PROP_NUMBERGROUPSEPARATOR)
        ReadLongArray .NumberGroupSizes, PROP_NUMBERGROUPSIZES, PropBag
        .NumberNegativePattern = PropBag.ReadProperty(PROP_NUMBERNEGATIVEPATTERN)
        .PercentDecimalDigits = PropBag.ReadProperty(PROP_PERCENTDECIMALDIGITS)
        .PercentDecimalSeparator = PropBag.ReadProperty(PROP_PERCENTDECIMALSEPARATOR)
        .PercentGroupSeparator = PropBag.ReadProperty(PROP_PERCENTGROUPSEPARATOR)
        ReadLongArray .PercentGroupSizes, PROP_PERCENTGROUPSIZES, PropBag
        .PercentNegativePattern = PropBag.ReadProperty(PROP_PERCENTNEGATIVEPATTERN)
        .PercentPositivePattern = PropBag.ReadProperty(PROP_PERCENTPOSITIVEPATTERN)
        .PercentSymbol = PropBag.ReadProperty(PROP_PERCENTSYMBOL)
        .PerMilleSymbol = PropBag.ReadProperty(PROP_PERMILLESYMBOL)
        .PositiveInfinitySymbol = PropBag.ReadProperty(PROP_POSITIVEINFINITYSYMBOL)
        .PositiveSign = PropBag.ReadProperty(PROP_POSITIVESIGN)
    End With
End Sub


Private Sub Class_Terminate()
    SAPtr(mNumber.digits) = 0
    CoTaskMemFree mPtrNumberBuffer
    CoTaskMemFree mBuffer.SA.pvData
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With mProps
        PropBag.WriteProperty PROP_CURRENCYDECIMALDIGITS, .CurrencyDecimalDigits
        PropBag.WriteProperty PROP_CURRENCYDECIMALSEPARATOR, .CurrencyDecimalSeparator
        PropBag.WriteProperty PROP_CURRENcYGROUPSEPARATOR, .CurrencyGroupSeparator
        WriteLongArray .CurrencyGroupSizes, PROP_CURRENCYGROUPSIZES, PropBag
        PropBag.WriteProperty PROP_CURRENCYNEGATIVEPATTERN, .CurrencyNegativePattern
        PropBag.WriteProperty PROP_CURRENCYPOSITIVEPATTERN, .CurrencyPositivePattern
        PropBag.WriteProperty PROP_CURRENCYSYMBOL, .CurrencySymbol
        PropBag.WriteProperty PROP_NANSYMBOL, .NaNSymbol
        PropBag.WriteProperty PROP_NEGATIVEINFINITYSYMBOL, .NegativeInfinitySymbol
        PropBag.WriteProperty PROP_NEGATIVESIGN, .NegativeSign
        PropBag.WriteProperty PROP_NUMBERDECIMALDIGITS, .NumberDecimalDigits
        PropBag.WriteProperty PROP_NUMBERDECIMALSEPARATOR, .NumberDecimalSeparator
        PropBag.WriteProperty PROP_NUMBERGROUPSEPARATOR, .NumberGroupSeparator
        WriteLongArray .NumberGroupSizes, PROP_NUMBERGROUPSIZES, PropBag
        PropBag.WriteProperty PROP_NUMBERNEGATIVEPATTERN, .NumberNegativePattern
        PropBag.WriteProperty PROP_PERCENTDECIMALDIGITS, .PercentDecimalDigits
        PropBag.WriteProperty PROP_PERCENTDECIMALSEPARATOR, .PercentDecimalSeparator
        PropBag.WriteProperty PROP_PERCENTGROUPSEPARATOR, .PercentGroupSeparator
        WriteLongArray .PercentGroupSizes, PROP_PERCENTGROUPSIZES, PropBag
        PropBag.WriteProperty PROP_PERCENTNEGATIVEPATTERN, .PercentNegativePattern
        PropBag.WriteProperty PROP_PERCENTPOSITIVEPATTERN, .PercentPositivePattern
        PropBag.WriteProperty PROP_PERCENTSYMBOL, .PercentSymbol
        PropBag.WriteProperty PROP_PERMILLESYMBOL, .PerMilleSymbol
        PropBag.WriteProperty PROP_POSITIVEINFINITYSYMBOL, .PositiveInfinitySymbol
        PropBag.WriteProperty PROP_POSITIVESIGN, .PositiveSign
    End With
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    cObject_Equals = Equals(Value)
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

Private Function IFormatProvider_GetFormat(ByVal FormatType As String) As Object
    Set IFormatProvider_GetFormat = GetFormat(FormatType)
End Function
