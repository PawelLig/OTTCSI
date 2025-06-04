Imports Mongoose.IDO
Imports Mongoose.IDO.Protocol

<IDOExtensionClass("ue_CustomerReplication")>
Public Class ue_CustomerReplication
    Inherits ExtensionClassBase
    <IDOMethod(MethodFlags.None, "Infobar")>
    Public Function ue_CustomerBankCodeReplVb(ByVal pCustNum As String,
                                              ByVal pCustSeq As String,
                                              ByVal pCurrCode As String,
                                              ByVal pDefBankCode As String,
                                              ByRef Infobar As String
                                             ) As Integer
        ue_CustomerBankCodeReplVb = 0

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
            Dim sFilter As String = "Name = 'OTT_CustBankCodeReplSites'"
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

                    'Update the Customer
                    Dim oUpdateCol As LoadCollectionResponseData
                    Dim sUpdateFilter As String
                    Dim oRequest As New UpdateCollectionRequestData
                    Dim oResponse As New UpdateCollectionResponseData
                    Dim oUpdateItem As New IDOUpdateItem

                    sUpdateFilter = "SiteRef = '" & ParsedSites(i) & "' AND CustNum = '" & pCustNum & "' AND CustSeq =" & pCustSeq

                    oUpdateCol = Me.Context.Commands.LoadCollection("ue_SLCustomerMsts",
                                                                    "SiteRef,
                                                                    CustNum,
                                                                    CustSeq",
                                                                    sUpdateFilter,
                                                                    "",
                                                                    1)
                    If oUpdateCol.Items.Count > 0 Then
                        oRequest = New UpdateCollectionRequestData("ue_SLCustomerMsts")
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
            ue_CustomerBankCodeReplVb = -1
            Infobar = ex.Message
            Exit Function
        End Try
    End Function
    <IDOMethod(MethodFlags.None, "Infobar")>
    Public Function ue_CustomerVATCodeReplVb(ByVal pCustNum As String,
                                             ByVal pCustSeq As String,
                                             ByVal pCustECCode As String,
                                             ByVal pCustVATCode As String,
                                             ByVal pCustCountry As String,
                                             ByRef Infobar As String
                                             ) As Integer
        ue_CustomerVATCodeReplVb = 0

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
            Dim sFilter As String = "Name = 'OTT_CustVATCodeReplSites'"
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

                    'Set the VAT Code for the Customer
                    If pCustSeq = "0" Then
                        If pCustVATCode = "NT" Then
                            sSiteVATCode = "NT"
                        ElseIf ParsedSites(i) = "STERLING" And (pCustCountry <> "CANADA" And pCustCountry <> "UNITED STATES") Then
                            sSiteVATCode = "EXPORT"
                        ElseIf ParsedSites(i) = "STERLING" And (pCustCountry = "CANADA" Or pCustCountry = "UNITED STATES") Then
                            sSiteVATCode = "EXTRNL"
                        ElseIf pCustECCode = sSiteCountryISO Then
                            sSiteVATCode = Nothing
                        ElseIf ParsedSites(i) <> "STERLING" And pCustECCode = Nothing Then
                            sSiteVATCode = "EXPORT"
                        Else
                            sSiteVATCode = pCustVATCode 'EU-I or EU-R if Service Ship To
                        End If
                    End If

                    If pCustSeq <> "0" Then
                        If pCustVATCode = "NT" Then
                            sSiteVATCode = "NT"
                        ElseIf ParsedSites(i) = "STERLING" And (pCustCountry <> "CANADA" And pCustCountry <> "UNITED STATES") Then
                            sSiteVATCode = "EXPORT"
                        ElseIf ParsedSites(i) = "STERLING" And (pCustCountry = "CANADA" Or pCustCountry = "UNITED STATES") Then
                            sSiteVATCode = "EXTRNL"
                        ElseIf pCustECCode = sSiteCountryISO Then
                            sSiteVATCode = Nothing
                        ElseIf ParsedSites(i) <> "STERLING" And pCustECCode = Nothing Then
                            sSiteVATCode = "EXPORT"
                        Else
                            sSiteVATCode = pCustVATCode 'EU-I or EU-R if Service Ship To
                        End If
                    End If

                    'Update the Customer
                    Dim oUpdateCol As LoadCollectionResponseData
                    Dim sUpdateFilter As String
                    Dim oRequest As New UpdateCollectionRequestData
                    Dim oResponse As New UpdateCollectionResponseData
                    Dim oUpdateItem As New IDOUpdateItem

                    sUpdateFilter = "SiteRef = '" & ParsedSites(i) & "' AND CustNum = '" & pCustNum & "' AND CustSeq =" & pCustSeq

                    oUpdateCol = Me.Context.Commands.LoadCollection("ue_SLCustomerMsts",
                                                                    "SiteRef,
                                                                     CustNum,
                                                                     CustSeq",
                                                                    sUpdateFilter,
                                                                    "",
                                                                    1)
                    If oUpdateCol.Items.Count > 0 Then
                        oRequest = New UpdateCollectionRequestData("ue_SLCustomerMsts")
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
            Infobar = ex.Message
            ue_CustomerVATCodeReplVb = -1
            Exit Function
        End Try
    End Function
End Class