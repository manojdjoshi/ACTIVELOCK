Option Strict Off
Option Explicit On
Friend Class clsShortSerial
	'*******************************************************************************
	' MODULE:       clsShortSerial based on clsOwnerRegistration originally
	' FILENAME:     clsShortSerial.cls
	' AUTHOR:       Phil Fresle
	' CREATED:      26-May-2001
	' COPYRIGHT:    Copyright 2001 Frez Systems Limited. All Rights Reserved.
	'
	' DESCRIPTION:
	' This class is designed to create a 16 character licence key based on the
	' name of the registered owner. The idea being that some person/company
	' purchases a licence for your product. You take their name and generate a
	' key for the application based on their name. Your application has a
	' registration form that requests the registered name and key. The key is
	' regenerated from the name and the key they enter is checked for equality.
	'
	' This method helps protect from pirating, as even if another user has a valid
	' registered user's name (other than their own name) and key combination, you
	' can ensure that the user's name appears on splash screens, title bars,
	' about forms, and more importantly all reports to embarass the pirateer
	' if they show their customer's the reports.
	'
	' An MD5 digest is made of the registered name after appending some characters
	' specific to our application (this might be the applications name
	' or simply some random combination of characters). The longer the string the
	' more difficult it will be for a hacker who knows our method to crack.
	'
	' MD5 returns a 16 byte digest in hex. We convert this back to bytes and obtain
	' a MOD 32 of each byte. We use the resultant number to lookup in our list
	' of valid characters. The 16 MD5 generated characters become the key.
	'
	' By supplying different application specific characters to be added before
	' generating the MD5 we can be fairly sure of uniqueness between applications.
	'
	' A possible enhancement to the routine might be to include licence type details
	' as extra characters that are then also used to generate the MD5. You would end
	' up with a longer key, but the key might then indicate that it was an upgrade
	' key, or for a specific version number, or perhaps a network rather than
	' standalone licence.
	'
	' This is 'free' software with the following restrictions:
	'
	' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
	' to use the source code in your own code, but you may not claim that you created
	' the sample code. It is expressly forbidden to sell or profit from this source code
	' other than by the knowledge gained or the enhanced value added by your own code.
	'
	' Use of this software is also done so at your own risk. The code is supplied as
	' is without warranty or guarantee of any kind.
	'
	' Should you wish to commission some derivative work based on this code provided
	' here, or any consultancy work, please do not hesitate to contact us.
	'
	' Web Site:  http://www.frez.co.uk
	' E-mail:    sales@frez.co.uk
	'
	' MODIFICATION HISTORY:
	' 1.0       27-May-2001
	'           Phil Fresle
	'           Initial Version
	'*******************************************************************************
	
	' 32 Numerics and alphas - we are missing I, O, S and Z not just because
	' we only want 32 characters but also because I could be mistaken for
	' the number 1, O for 0, S for 5 and Z for 2.
	Private Const VALID_CHARS As String = "0123456789" 'ABCDEFGHJKLMNPQRTUVWXY"
	
	Private Const RANDOM_LOWER As Integer = 0
	Private Const RANDOM_UPPER As Integer = 31
	
	'*******************************************************************************
	' Function GenerateKey
	'
	' PARAMETERS:
	' (In/Out) - sAppNameVersionPassword - String - Application specific composite string
	' (In/Out) - sHDDfirmwareSerial      - String - HDD firmware serial number
	'
	' RETURN VALUE:
	' String - The key
	'
	' DESCRIPTION:
	' Generates a short serial number by taking the app name, adding the app version
	' number and the app password, creating an MD5 digest, and using the digest to select
	' the 8 characters for our serial.
	'*******************************************************************************
	Public Function GenerateKey(ByRef sAppNameVersionPassword As String, ByRef sHDDfirmwareSerial As String) As String
		Dim lChar As Integer
        Dim lCount As Integer
        Dim sMD5 As String
        Dim sKey As String

        ' We now get an MD5 of the combined string.
        ' We use an upper case version of the registered name as we don't want to
        ' penalise the user for our mistaken capitalisation of a name like
        ' MacGewan, or for a company INC rather than Inc. The user will need to
        ' be warned about enterign the name exactly as specified, as a difference
        ' in spelling or punctuation would be critical.
        '
        ' The application name+version+password should be different for each application to
        ' ensure that a key that is valid for one application is not valid for another
        ' application.
        ' Try to make the application characters long to help prevent cracking.
        sMD5 = ComputeHash(UCase(sAppNameVersionPassword) & sHDDfirmwareSerial)
		
		' We now take each byte-pair from the MD5, convert it back to a byte
		' value from the hex code, do a mod 10, and then select the appropriate
		' character from our VALID_CHARS
		' Valid characters are numbers from 0-9
		sKey = ""
		For lCount = 1 To 8 ' doing only 8-digit derial number for ease of use
			lChar = CInt("&H" & Mid(sMD5, (lCount * 2) - 1, 2)) Mod 10 'doing only 10 digits instead of say 32
			sKey = sKey & Mid(VALID_CHARS, lChar + 1, 1)
		Next 
		
		GenerateKey = sKey
	End Function
	
	'*******************************************************************************
	' IsKeyOK (FUNCTION)
	'
	' PARAMETERS:
	' (In/Out) - sKey      - String - Key to check
	' (In/Out) - sAppChars - String - Application specific characters used in
	'                                 generating the key.
	'
	' RETURN VALUE:
	' Boolean - True if valid
	'
	' DESCRIPTION:
	' Takes the key, recalculates the MD5 part and tests for equality.
	'*******************************************************************************
	Public Function IsKeyOK(ByRef sKey As String, ByRef sAppNameVersionPassword As String, ByRef sHDDfirmwareSerial As String) As Boolean
		Dim Hash As Object
		
		Dim lChar As Integer
		Dim lCount As Integer
		'Dim oMD5        As CMD5
		Dim sMD5 As String
		Dim sTestKey As String
		
		' Recalculate the MD5 digest
        sMD5 = ComputeHash(UCase(sAppNameVersionPassword) & sHDDfirmwareSerial)

		' We now take each byte-pair from the MD5, convert it back to a byte
		' value from the hex code, do a MOD 32, and then select the appropriate
		' character from our VALID_CHARS
		sTestKey = ""
		For lCount = 1 To 8
			lChar = CInt("&H" & Mid(sMD5, (lCount * 2) - 1, 2)) Mod 10 '32
			sTestKey = sTestKey & Mid(VALID_CHARS, lChar + 1, 1)
		Next 
		
		' Check for equality
		IsKeyOK = (sTestKey = sKey)
	End Function
End Class