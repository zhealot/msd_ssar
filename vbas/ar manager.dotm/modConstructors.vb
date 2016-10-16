Attribute VB_Name = "modConstructors"
'===================================================================================================================================
' Module:       modConstructors
' Purpose:      Because the RDA and SSAR addins cannot create new class instances of classes within this addin we expose a public
'               constructor function for each class that needs to be created by the RDA or SSAR addin.
'
' Author:       Peter Hewett - Inner Word Limited (innerword@xnet.co.nz)
' Copyright:    Ministry of Social Development (MSD) ©2016 All rights reserved.
' Contact       Inner Word Limited
' details:      134 Kahu Road
'               Paremata
'               Porirua City
'               5024
'               T: +64 4 233 2124
'               M: +64 21 213 5063
'               E: innerword@xnet.co.nz
'
' History:      31/05/2016    1.  Created.
'===================================================================================================================================
Option Explicit


Public Function NewDocumentProtection() As DocumentProtection
    Set NewDocumentProtection = New DocumentProtection
End Function ' NewDocumentProtection
