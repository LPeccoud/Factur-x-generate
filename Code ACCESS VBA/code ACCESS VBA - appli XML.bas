Option Compare Database
Option Explicit

Sub ExportFactureXML(PDF_Facture As String, FacRef As Long) 'chemin du PDF de la facture , ID de la facture dans la BD
        Dim Doc As MSXML2.DOMDocument60
        Dim root As IXMLDOMElement, node As IXMLDOMNode, block As IXMLDOMNode, sousBlock As IXMLDOMNode, PostalBock As IXMLDOMNode, newNode As IXMLDOMNode
        Dim lineModel As IXMLDOMNode, lineClone As IXMLDOMNode, agreementNode As IXMLDOMNode
        Dim NumFac As String, Design As String, Descro As String, Total As String, avance As String
        Dim solde As String, DateExe As String, DateFac As String, Echéance As String, Ref As String
        Dim PrixUnitaire As String, Quantité As String, Montant As String
        Dim Client As String, CP As String, Ville As String, Mail As String, SIREN As String, Adresse_1 As String, Adresse_2 As String
        Dim n As Integer, f As Integer
        Dim rs As DAO.Recordset, qdf As DAO.QueryDef
        Dim chemin As String, fichierXml As String, fichierLog As String, script As String, pathXml As String
        Dim sh As WshShell, exec As WshExec, cmd As String, Sortie As String, erreurs As String
    
        chemin = "C:\Users\XXX\Documents\Factur-X\" 'chemin à personnaliser
        fichierXml = chemin & "factur-X-en16931_MicroEI.xml" ' modèle XML à utiliser
        fichierLog = chemin & "Factur-X_Log.txt" ' fichier Log de retour
        pathXml = chemin & "Factur-X_Temp.xml" ' fichier XML à insérer dans le PDF
        script = chemin & "FacturX_Insert.py" ' Script Python d'insertion
   
        '=========================
        ' FACTURE
        '=========================

        Set qdf = CurrentDb.QueryDefs("Requête_FactureX_facture")
        qdf.Parameters("FacRef") = FacRef

        Set rs = qdf.OpenRecordset
        '    root = root.DocumentElement
        If Not rs.EOF Then
                With rs
                        Total = Replace(Format(Nz(!Total, 0), "0.00"), ",", ".")
                        avance = Replace(Format(Nz(!avance, 0), "0.00"), ",", ".")
                        solde = Replace(Format(Nz(!Total, 0) - Nz(!avance, 0), "0.00"), ",", ".")
                        DateExe = Format(Nz(!Date_exe, !DateFac), "yyyymmdd")
                        Echéance = Format(Nz(!échéance, !DateFac + 5), "yyyymmdd")
                        DateFac = Format(!DateFac, "yyyymmdd")
                        NumFac = !NumFac
                        Ref = Nz(!Reférence, "")
                        .Close
                End With
                '=========================
                ' CLIENT
                '=========================
        End If
        Set qdf = CurrentDb.QueryDefs("Requête_FactureX_client")
        qdf.Parameters("FacRef") = FacRef

        Set rs = qdf.OpenRecordset
        If Not rs.EOF Then
                With rs
                        Client = !Client
                        SIREN = Left(Nz(!SIRET, ""), 9)
                        Mail = Nz(!Email, "")
                        If Mail <> "" Then Mail = Split(Mail, "#")(0)
                        CP = !CP
                        Ville = !Ville
                        Adresse_1 = ![Adresse 1]
                        Adresse_2 = Nz(![Adresse 2], "")
                        .Close
                End With
        End If
        '=========================
        ' CHARGEMENT XML
        '=========================

        Set Doc = New MSXML2.DOMDocument60
        Doc.async = False
        Doc.validateOnParse = False
        Doc.SetProperty "SelectionNamespaces", _
                "xmlns:rsm='urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100' " & _
                "xmlns:ram='urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100' " & _
                "xmlns:udt='urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100'"

        If Not Doc.Load(fichierXml) Then
'                MsgBox root.parseError.reason
                Exit Sub
        End If
        
        Set root = Doc.SelectSingleNode("./rsm:CrossIndustryInvoice")
        Set block = root.SelectSingleNode("./rsm:ExchangedDocument")
        With block
                .SelectSingleNode("./ram:ID").text = NumFac
                .SelectSingleNode("./ram:IssueDateTime/udt:DateTimeString").text = DateFac
        End With
    
        Set block = root.SelectSingleNode("./rsm:SupplyChainTradeTransaction/ram:ApplicableHeaderTradeAgreement")
        With block
                Set sousBlock = .SelectSingleNode("./ram:BuyerTradeParty")
                With sousBlock
                        .SelectSingleNode("./ram:Name").text = Client
                        ' SIREN
                        Set node = .SelectSingleNode("./ram:SpecifiedLegalOrganization/ram:ID")
                        setNode node, SIREN
            
                        Set PostalBock = .SelectSingleNode("./ram:PostalTradeAddress")
                        With PostalBock
                                .SelectSingleNode("./ram:PostcodeCode").text = CP
                                .SelectSingleNode("./ram:CityName").text = Ville
                                .SelectSingleNode("./ram:CountryID").text = "FR"
                                ' Adresse 1
                                Set node = .SelectSingleNode("./ram:LineOne")
                                setNode node, Adresse_1
                                ' Adresse 2
                                Set node = .SelectSingleNode("./ram:LineTwo")
                                setNode node, Adresse_2
                        End With
            
                        Set node = .SelectSingleNode("./ram:URIUniversalCommunication/ram:URIID")
                        setNode node, SIREN
                End With
            
                Set node = .SelectSingleNode("./ram:BuyerOrderReferencedDocument/ram:IssuerAssignedID")
                setNode node, Ref
        End With
        
        root.SelectSingleNode("./rsm:SupplyChainTradeTransaction/ram:ApplicableHeaderTradeDelivery/ram:ActualDeliverySupplyChainEvent/ram:OccurrenceDateTime/udt:DateTimeString").text = DateExe
    
        Set block = root.SelectSingleNode("./rsm:SupplyChainTradeTransaction/ram:ApplicableHeaderTradeSettlement")
        With block
                .SelectSingleNode("./ram:ApplicableTradeTax/ram:BasisAmount").text = Total
                .SelectSingleNode("./ram:BillingSpecifiedPeriod/ram:EndDateTime/udt:DateTimeString").text = DateExe
                .SelectSingleNode("./ram:SpecifiedTradePaymentTerms/ram:DueDateDateTime/udt:DateTimeString").text = Echéance
                Set sousBlock = .SelectSingleNode("./ram:SpecifiedTradeSettlementHeaderMonetarySummation")
                With sousBlock
                        .SelectSingleNode("./ram:LineTotalAmount").text = Total
                        .SelectSingleNode("./ram:TaxBasisTotalAmount").text = Total
                        .SelectSingleNode("./ram:GrandTotalAmount").text = Total
                        .SelectSingleNode("./ram:TotalPrepaidAmount").text = avance
                        .SelectSingleNode("./ram:DuePayableAmount").text = solde
                End With
        End With


        '=========================
        ' LIGNES FACTURE
        '=========================
        Set block = root.SelectSingleNode("./rsm:SupplyChainTradeTransaction")
        Set lineModel = block.SelectSingleNode("./ram:IncludedSupplyChainTradeLineItem")

        If lineModel Is Nothing Then
                MsgBox "Modèle ligne introuvable"
                Exit Sub
        End If

        Set qdf = CurrentDb.QueryDefs("Requête_FactureX_lignes")
        qdf.Parameters("FacRef") = FacRef

        Set rs = qdf.OpenRecordset
        With rs
                Do While Not .EOF
                        PrixUnitaire = Replace(Format(Nz(!PrixUnitaire, 0), "0.00"), ",", ".")
                        Quantité = Replace(Format(Nz(!Quantité, 0), "0.00"), ",", ".")
                        Montant = Replace(Format(Nz(!Montant, 0), "0.00"), ",", ".")
                        Descro = XmlSafe(!Descro)
                        Design = XmlSafe(!Design)
                        If Design = "" Then
                                Design = Descro
                                Descro = ""
                        End If
                        Set lineClone = lineModel.CloneNode(True)
        
                        lineClone.SelectSingleNode("./ram:AssociatedDocumentLineDocument/ram:LineID").text = !§
                        Set sousBlock = lineClone.SelectSingleNode("./ram:SpecifiedTradeProduct")
                        sousBlock.SelectSingleNode("./ram:Name").text = Design
                        setNode sousBlock.SelectSingleNode("./ram:Description"), Descro
        
                        Set sousBlock = lineClone.SelectSingleNode("./ram:SpecifiedLineTradeAgreement")
                        sousBlock.SelectSingleNode("./ram:GrossPriceProductTradePrice/ram:ChargeAmount").text = PrixUnitaire
                        sousBlock.SelectSingleNode("./ram:GrossPriceProductTradePrice/ram:BasisQuantity").text = Quantité
                        sousBlock.SelectSingleNode("./ram:NetPriceProductTradePrice/ram:ChargeAmount").text = PrixUnitaire
                        sousBlock.SelectSingleNode("./ram:NetPriceProductTradePrice/ram:BasisQuantity").text = Quantité
        
                        Set sousBlock = lineClone.SelectSingleNode("./ram:SpecifiedLineTradeDelivery")
                        sousBlock.SelectSingleNode("./ram:BilledQuantity").text = Quantité

                        Set sousBlock = lineClone.SelectSingleNode("./ram:SpecifiedLineTradeSettlement/ram:SpecifiedTradeSettlementLineMonetarySummation")
                        sousBlock.SelectSingleNode("./ram:LineTotalAmount").text = Montant

                        Set agreementNode = block.SelectSingleNode("./ram:ApplicableHeaderTradeAgreement")
                        block.InsertBefore lineClone, agreementNode
                        .MoveNext
                Loop
                .Close
        End With


        ' supprimer ligne modèle
        block.RemoveChild lineModel
    
        '=========================
        ' SAUVEGARDE
        '=========================
        Doc.Save pathXml
        
        Set sh = New WshShell
        cmd = "py -3.12  """ & script & """ """ & pathXml & """ """ & PDF_Facture & """"
        Set exec = sh.exec(cmd)
        ' attendre la fin
        Do While exec.Status = 0
                DoEvents
        Loop

        Sortie = exec.StdOut.ReadAll
        erreurs = exec.StdErr.ReadAll

        If exec.ExitCode = 0 Then
                MsgBox "Factur-X réalisée avec succès sur " & PDF_Facture, vbInformation
        Else
                MsgBox _
                        "Erreur Factur-X :" & vbCrLf & vbCrLf & erreurs & vbCrLf & Sortie, vbCritical
        End If
        f = FreeFile()
        Open fichierLog For Output As #f
        Print #f, cmd
        Print #f, erreurs
        Close #f
        Debug.Print cmd
        Debug.Print erreurs
        Application.FollowHyperlink PDF_Facture
        Set rs = Nothing: Set qdf = Nothing
End Sub

Sub setNode(node As IXMLDOMNode, valeur As Variant)
        Dim parent As IXMLDOMNode, n As Integer, child As IXMLDOMNode
        valeur = XmlSafe(Nz(valeur, ""))
        If valeur = "" Then
        ' suppression récursive des parents s'ils deviennent vides
                Do
                        Set parent = node.ParentNode
                        parent.RemoveChild node
                        For Each child In parent.ChildNodes
                                If child.NodeType = NODE_ELEMENT Then Exit Do
                        Next
                        Set node = parent
                Loop
        Else
                node.text = valeur
        End If
End Sub

Function XmlSafe(txt As String) As String
        txt = Replace(txt, "&", "&amp;")
        txt = Replace(txt, "<", "&lt;")
        txt = Replace(txt, ">", "&gt;")
        txt = Replace(txt, """", "&quot;")
        txt = Replace(txt, "'", "&apos;")
        XmlSafe = Trim(txt)
End Function
