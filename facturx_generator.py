import json
import re
from datetime import datetime
from lxml import etree
from facturx import generate_from_file


# =========================
# CONFIG ENTREPRISE
# =========================

SELLER = {
    "name": "MA SOCIETE",
    "siret": "12345678900012",
    "vat": "FR12345678901",
    "iban": "FR7612345987650123456789014",
    "bic": "AGRIFRPPXXX"
}


# =========================
# UTILS
# =========================

def clean(s):
    return re.sub(r"<.*?>", "", str(s or "")).strip()


def format_date(d):
    return datetime.strptime(d, "%d/%m/%Y %H:%M:%S").strftime("%Y%m%d")

# =========================
# XML BUILD
# =========================

def build_xml(data):

    NS = {
        "rsm": "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100",
        "ram": "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100",
        "udt": "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100"
    }

    root = etree.Element("{%s}CrossIndustryInvoice" % NS["rsm"], nsmap=NS)

    # CONTEXT
    ctx = etree.SubElement(root, "{%s}ExchangedDocumentContext" % NS["rsm"])
    param = etree.SubElement(ctx, "{%s}GuidelineSpecifiedDocumentContextParameter" % NS["ram"])
    etree.SubElement(param, "{%s}ID" % NS["ram"]).text = "urn:cen.eu:en16931:2017"

    # =========================
    # DOCUMENT
    # =========================
    doc = etree.SubElement(root, "{%s}ExchangedDocument" % NS["rsm"])

    etree.SubElement(doc, "{%s}ID" % NS["ram"]).text = data["facture"]["numero"]
    etree.SubElement(doc, "{%s}TypeCode" % NS["ram"]).text = "380"

    # 🔥 AJOUT OBLIGATOIRE EN16931 (FIX ERREUR COMFORT)
    # etree.SubElement(doc, "{%s}BuyerReference" % NS["ram"]).text = "CLIENT"

    issue = etree.SubElement(doc, "{%s}IssueDateTime" % NS["ram"])
    dt = etree.SubElement(issue, "{%s}DateTimeString" % NS["udt"], format="102")
    dt.text = format_date(data["facture"]["date"])

    # TRANSACTION
    trade = etree.SubElement(root, "{%s}SupplyChainTradeTransaction" % NS["rsm"])



    # =========================
    # LIGNES
    # =========================

    total_ht = 0
    tva_map = {}
    total_tva = 0

    for i, l in enumerate(data["lignes"], 1):

        qty = float(l["quantite"])
        price = float(l["prix"])
        tva = float(l.get("tva", 20))

        total = round(qty * price, 2)
        total_tva += round(total*tva/100,2)
        total_ht += total

        tva_map.setdefault(tva, 0)
        tva_map[tva] += total

        line = etree.SubElement(trade, "{%s}IncludedSupplyChainTradeLineItem" % NS["ram"])

        etree.SubElement(
            etree.SubElement(line, "{%s}AssociatedDocumentLineDocument" % NS["ram"]),
            "{%s}LineID" % NS["ram"]
        ).text = str(i)

        etree.SubElement(
            etree.SubElement(line, "{%s}SpecifiedTradeProduct" % NS["ram"]),
            "{%s}Name" % NS["ram"]
        ).text = clean(l["designation"])

        price_node = etree.SubElement(
            etree.SubElement(line, "{%s}SpecifiedLineTradeAgreement" % NS["ram"]),
            "{%s}NetPriceProductTradePrice" % NS["ram"]
        )
        etree.SubElement(price_node, "{%s}ChargeAmount" % NS["ram"]).text = str(price)

        etree.SubElement(
            etree.SubElement(line, "{%s}SpecifiedLineTradeDelivery" % NS["ram"]),
            "{%s}BilledQuantity" % NS["ram"],
            unitCode="C62"
        ).text = str(qty)

        # TVA ligne
        line_settle = etree.SubElement(line, "{%s}SpecifiedLineTradeSettlement" % NS["ram"])
        tax_line = etree.SubElement(line_settle, "{%s}ApplicableTradeTax" % NS["ram"])
        etree.SubElement(tax_line, "{%s}TypeCode" % NS["ram"]).text = "VAT"
        etree.SubElement(tax_line, "{%s}CategoryCode" % NS["ram"]).text = "S"
        etree.SubElement(tax_line, "{%s}RateApplicablePercent" % NS["ram"]).text = str(tva)

        
        mon_sum = etree.SubElement(line_settle, "{%s}SpecifiedTradeSettlementLineMonetarySummation" % NS["ram"])
        etree.SubElement(mon_sum, "{%s}LineTotalAmount" % NS["ram"]).text = str(total)
    # =========================
    # PARTIES
    # =========================
    print(total_tva)
    agr = etree.SubElement(trade, "{%s}ApplicableHeaderTradeAgreement" % NS["ram"])

    seller = etree.SubElement(agr, "{%s}SellerTradeParty" % NS["ram"])
    etree.SubElement(seller, "{%s}Name" % NS["ram"]).text = SELLER["name"]

    # SIRET
    legal = etree.SubElement(seller, "{%s}SpecifiedLegalOrganization" % NS["ram"])
    etree.SubElement(legal, "{%s}ID" % NS["ram"], schemeID="0002").text = SELLER["siret"]

    addr_seller = etree.SubElement(seller, "{%s}PostalTradeAddress" % NS["ram"])
    etree.SubElement(addr_seller, "{%s}PostcodeCode" % NS["ram"]).text = "75000"
    etree.SubElement(addr_seller, "{%s}LineOne" % NS["ram"]).text = "Adresse société"
    etree.SubElement(addr_seller, "{%s}CityName" % NS["ram"]).text = "PARIS"
    etree.SubElement(addr_seller, "{%s}CountryID" % NS["ram"]).text = "FR"

    # TVA vendeur
    taxreg = etree.SubElement(seller, "{%s}SpecifiedTaxRegistration" % NS["ram"])
    etree.SubElement(taxreg, "{%s}ID" % NS["ram"], schemeID="VA").text = SELLER["vat"]

    # CLIENT
    
    buyer = etree.SubElement(agr, "{%s}BuyerTradeParty" % NS["ram"])
    etree.SubElement(buyer, "{%s}Name" % NS["ram"]).text = clean(data["client"]["nom"])

    legal_buyer = etree.SubElement(buyer, "{%s}SpecifiedLegalOrganization" % NS["ram"])
    etree.SubElement(legal_buyer, "{%s}ID" % NS["ram"], schemeID="0002").text = "36573887100073"



    addr = etree.SubElement(buyer, "{%s}PostalTradeAddress" % NS["ram"])
    etree.SubElement(addr, "{%s}PostcodeCode" % NS["ram"]).text = data["client"]["cp"]
    etree.SubElement(addr, "{%s}LineOne" % NS["ram"]).text = clean(data["client"]["adresse1"])
    # etree.SubElement(addr, "{%s}LineTwo" % NS["ram"]).text = clean(data["client"]["adresse2"])
    etree.SubElement(addr, "{%s}CityName" % NS["ram"]).text = data["client"]["ville"]
    etree.SubElement(addr, "{%s}CountryID" % NS["ram"]).text = "FR"

    # DELIVERY
    delivery = etree.SubElement(trade, "{%s}ApplicableHeaderTradeDelivery" % NS["ram"])
    event = etree.SubElement(delivery, "{%s}ActualDeliverySupplyChainEvent" % NS["ram"])
    date = etree.SubElement(event, "{%s}OccurrenceDateTime" % NS["ram"])

    dt = etree.SubElement(date, "{%s}DateTimeString" % NS["udt"], format="102")
    dt.text = format_date(data["facture"]["date"])


    # =========================
    # SETTLEMENT
    # =========================

    settlement = etree.SubElement(trade, "{%s}ApplicableHeaderTradeSettlement" % NS["ram"])
    # etree.SubElement(settlement, "{%s}TaxCurrencyCode" % NS["ram"]).text = "EUR"
    etree.SubElement(settlement, "{%s}InvoiceCurrencyCode" % NS["ram"]).text = "EUR"

    payment = etree.SubElement(settlement, "{%s}SpecifiedTradeSettlementPaymentMeans" % NS["ram"])
    etree.SubElement(payment, "{%s}TypeCode" % NS["ram"]).text = "42"

    etree.SubElement(payment, "{%s}Information" % NS["ram"]).text = "Virement bancaire"

    # PAIEMENT (IBAN/BIC)
    account = etree.SubElement(payment, "{%s}PayeePartyCreditorFinancialAccount" % NS["ram"])
    etree.SubElement(account, "{%s}IBANID" % NS["ram"]).text = SELLER["iban"]


    # totaux de TVA par type
    total_tva = 0

    for taux, base in tva_map.items():
        tva_amount = round(base * taux / 100, 2)
        total_tva += tva_amount

        tax = etree.SubElement(settlement, "{%s}ApplicableTradeTax" % NS["ram"])
        etree.SubElement(tax, "{%s}CalculatedAmount" % NS["ram"]).text = str(total_tva)
        etree.SubElement(tax, "{%s}TypeCode" % NS["ram"]).text = "VAT"
        # etree.SubElement(tax, "{%s}ExemptionReason" % NS["ram"]).text = "VAT"
        etree.SubElement(tax, "{%s}BasisAmount" % NS["ram"]).text = str(total_ht)
        etree.SubElement(tax, "{%s}CategoryCode" % NS["ram"]).text = "S"
        # etree.SubElement(tax, "{%s}ExemptionReasonCode" % NS["ram"]).text = "VAT"
        etree.SubElement(tax, "{%s}RateApplicablePercent" % NS["ram"]).text = str(taux)

    # CONDITIONS
    terms = etree.SubElement(settlement, "{%s}SpecifiedTradePaymentTerms" % NS["ram"])
    etree.SubElement(terms, "{%s}Description" % NS["ram"]).text = "Paiement a 30 jours"

    total_ttc = round(total_ht + total_tva, 2)

    summary = etree.SubElement(settlement, "{%s}SpecifiedTradeSettlementHeaderMonetarySummation" % NS["ram"])
    etree.SubElement(summary, "{%s}LineTotalAmount" % NS["ram"]).text = str(total_ht)
    etree.SubElement(summary, "{%s}TaxBasisTotalAmount" % NS["ram"]).text = str(total_ht)
    tax_total = etree.SubElement(summary, "{%s}TaxTotalAmount" % NS["ram"])
    tax_total.text = str(total_tva)
    tax_total.set("currencyID", "EUR")
    etree.SubElement(summary, "{%s}GrandTotalAmount" % NS["ram"]).text = str(total_ttc)
    etree.SubElement(summary, "{%s}DuePayableAmount" % NS["ram"]).text = str(total_ttc)

    return etree.tostring(root, encoding="UTF-8")


# =========================
# MAIN
# =========================

def main():

    base = "C:/Users/Nemo/Documents/Entreprise/BD_Entreprise/"
    pdf = base + "facture.pdf"
    data = json.load(open(base + "facture.json", encoding="utf-8-sig"))

    xml = build_xml(data)
    # with open(base + "xml.txt", "r", encoding="utf-8-sig") as f:
    #     xml = f.read()
    generate_from_file(
        pdf,
        xml,
        flavor="factur-x",
        level="en16931:2017",
        output_pdf_file=pdf.replace(".pdf", "_facturx.pdf")
    )

    print("FACTUR-X PRODUCTION 2026 READY")


if __name__ == "__main__":
    main()