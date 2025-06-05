Imports System.ComponentModel
Imports Mongoose.IDO
Imports Mongoose.IDO.Metadata
Imports Mongoose.IDO.Protocol

<IDOExtensionClass("ue_VendorReplication")>
Public Class ue_VendorReplication
    Inherits ExtensionClassBase
    <IDOMethod(MethodFlags.None, "Infobar")>
    Public Function ue_VendorBankCodeReplVb(ByVal pVendNum As String,
                                            ByVal pCurrCode As String,
                                            ByVal pDefBankCode As String,
                                            ByRef Infobar As String
                                           ) As Integer

        ue_VendorBankCodeReplVb = 0

        Try
            'Get And Parse Event Global Constant
            Dim sSites As String 'GC
            Dim ParsedSites() As String 'Parsed Sites

            Dim oSiteCol As LoadCollectionResponseData
            Dim sSiteFilter As String
            Dim sSiteCountry As String
            Dim sSiteCountryISO As String

            Dim sSiteBankCode As String

            Dim oCol As LoadCollectionResponseData
            Dim sFilter As String = "Name = 'OTT_VendBankCodeReplSites'"
            oCol = Me.Context.Commands.LoadCollection("EventGlobalConstants", "Value", sFilter, "", 1)

            If oCol.Items.Count > 0 Then
                sSites = oCol.Item(0, "Value").Value
            End If

            ParsedSites = sSites.Split(";"c)

            'Loop on all sites in the Event Global Constant
            If ParsedSites.Length > 0 Then
                For i As Integer = 0 To ParsedSites.Length - 1
                    'Select site country
                    sSiteFilter = "SiteRef ='" & ParsedSites(i) & "' AND ParmKey = 0"
                    oSiteCol = Me.Context.Commands.LoadCollection("SLParmsAlls", "Country", sSiteFilter, "", 1)

                    If oSiteCol.Items.Count > 0 Then
                        sSiteCountry = oSiteCol.Item(0, "Country").Value
                    End If

                    'Get Site Country ISO Code
                    sSiteFilter = "Country ='" & sSiteCountry & "'"
                    oSiteCol = Me.Context.Commands.LoadCollection("SLCountries", "ISOCountryCode", sSiteFilter, "", 1)
                    If oSiteCol.Items.Count > 0 Then
                        sSiteCountryISO = oSiteCol.Item(0, "ISOCountryCode").Value
                    End If

                    'Set the Bank Code
                    If sSiteCountryISO = "NL" And pCurrCode = "USD" Then
                        sSiteBankCode = "DB3"
                    ElseIf sSiteCountryISO = "DE" Then
                        sSiteBankCode = "DB1"
                    ElseIf ParsedSites(i) = "STERLING" Then
                        sSiteBankCode = "BA1"
                    Else
                        sSiteBankCode = pDefBankCode
                    End If

                    'Update the Vendor
                    Dim oUpdateCol As LoadCollectionResponseData
                    Dim sUpdateFilter As String
                    Dim oRequest As New UpdateCollectionRequestData
                    Dim oResponse As New UpdateCollectionResponseData
                    Dim oUpdateItem As New IDOUpdateItem

                    sUpdateFilter = "SiteRef = '" & ParsedSites(i) & "' AND VendNum = '" & pVendNum & "'"

                    oUpdateCol = Me.Context.Commands.LoadCollection("ue_SLVendorMsts",
                                                                    "SiteRef,
                                                                    VendNum",
                                                                    sUpdateFilter,
                                                                    "",
                                                                    1)
                    If oUpdateCol.Items.Count > 0 Then
                        oRequest = New UpdateCollectionRequestData("ue_SLVendorMsts")
                        oRequest.RefreshAfterUpdate = True
                        oUpdateItem.ItemID = oUpdateCol.Items(0).ItemID
                        oUpdateItem.Action = UpdateAction.Update
                        oUpdateItem.Properties.Add("BankCode", sSiteBankCode)
                        oRequest.Items.Add(oUpdateItem)
                        oResponse = Me.Context.Commands.UpdateCollection(oRequest)
                    End If
                Next i
            End If
        Catch ex As Exception
            ue_VendorBankCodeReplVb = -1
            Infobar = ex.Message
            Exit Function
        End Try
    End Function
    <IDOMethod(MethodFlags.None, "Infobar")>
    Public Function ue_VendorVATCodeReplVb(ByVal pVendNum As String,
                                           ByVal pVendECCode As String,
                                           ByVal pVendVATCode As String,
                                           ByRef Infobar As String
                                           ) As Integer

        ue_VendorVATCodeReplVb = 0

        Try
            'Get And Parse Event Global Constant
            Dim sSites As String 'GC
            Dim ParsedSites() As String 'Parsed Sites

            Dim oSiteCol As LoadCollectionResponseData
            Dim sSiteFilter As String
            Dim sSiteCountry As String
            Dim sSiteCountryISO As String

            Dim sSiteVATCode As String

            Dim oCol As LoadCollectionResponseData
            Dim sFilter As String = "Name = 'OTT_VendVATCodeReplSites'"
            oCol = Me.Context.Commands.LoadCollection("EventGlobalConstants", "Value", sFilter, "", 1)

            If oCol.Items.Count > 0 Then
                sSites = oCol.Item(0, "Value").Value
            End If

            ParsedSites = sSites.Split(";"c)

            'Loop on all sites in the Event Global Constant
            If ParsedSites.Length > 0 Then
                For i As Integer = 0 To ParsedSites.Length - 1
                    'Select site country
                    sSiteFilter = "SiteRef ='" & ParsedSites(i) & "' AND ParmKey = 0"
                    oSiteCol = Me.Context.Commands.LoadCollection("SLParmsAlls", "Country", sSiteFilter, "", 1)

                    If oSiteCol.Items.Count > 0 Then
                        sSiteCountry = oSiteCol.Item(0, "Country").Value
                    End If

                    'Get Site Country ISO Code
                    sSiteFilter = "Country ='" & sSiteCountry & "'"
                    oSiteCol = Me.Context.Commands.LoadCollection("SLCountries", "ISOCountryCode", sSiteFilter, "", 1)
                    If oSiteCol.Items.Count > 0 Then
                        sSiteCountryISO = oSiteCol.Item(0, "ISOCountryCode").Value
                    End If

                    'Set the VAT Code
                    If pVendVATCode = "NT" Then
                        sSiteVATCode = "NT"
                    ElseIf ParsedSites(i) = "STERLING" Then
                        sSiteVATCode = "NA"
                    ElseIf pVendECCode = sSiteCountryISO Then
                        sSiteVATCode = Nothing
                    ElseIf ParsedSites(i) <> "STERLING" And pVendECCode = Nothing Then
                        sSiteVATCode = "IMPORT"
                    Else
                        sSiteVATCode = "EU"
                    End If

                    'Update the Vendor
                    Dim oUpdateCol As LoadCollectionResponseData
                    Dim sUpdateFilter As String
                    Dim oRequest As New UpdateCollectionRequestData
                    Dim oResponse As New UpdateCollectionResponseData
                    Dim oUpdateItem As New IDOUpdateItem

                    sUpdateFilter = "SiteRef = '" & ParsedSites(i) & "' AND VendNum = '" & pVendNum & "'"

                    oUpdateCol = Me.Context.Commands.LoadCollection("ue_SLVendorMsts",
                                                                    "SiteRef,
                                                                    VendNum",
                                                                    sUpdateFilter,
                                                                    "",
                                                                    1)
                    If oUpdateCol.Items.Count > 0 Then
                        oRequest = New UpdateCollectionRequestData("ue_SLVendorMsts")
                        oRequest.RefreshAfterUpdate = True
                        oUpdateItem.ItemID = oUpdateCol.Items(0).ItemID
                        oUpdateItem.Action = UpdateAction.Update
                        oUpdateItem.Properties.Add("TaxCode1", sSiteVATCode)
                        oRequest.Items.Add(oUpdateItem)
                        oResponse = Me.Context.Commands.UpdateCollection(oRequest)
                    End If
                Next i
            End If
        Catch ex As Exception
            ue_VendorVATCodeReplVb = -1
            Infobar = ex.Message
            Exit Function
        End Try
    End Function
End Class