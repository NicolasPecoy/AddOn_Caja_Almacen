using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOn_Caja.Clases
{

    public class Pago
    {
        public string odatametadata { get; set; }
        public int DocNum { get; set; }
        public string DocType { get; set; }
        public string HandWritten { get; set; }
        public string Printed { get; set; }
        public string DocDate { get; set; }
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public object Address { get; set; }
        public object CashAccount { get; set; }
        public string DocCurrency { get; set; }
        public float CashSum { get; set; }
        public object CheckAccount { get; set; }
        public object TransferAccount { get; set; }
        public float TransferSum { get; set; }
        public object TransferDate { get; set; }
        public object TransferReference { get; set; }
        public string LocalCurrency { get; set; }
        public float DocRate { get; set; }
        public string Reference1 { get; set; }
        public object Reference2 { get; set; }
        public object CounterReference { get; set; }
        public string Remarks { get; set; }
        public string JournalRemarks { get; set; }
        public string SplitTransaction { get; set; }
        public object ContactPersonCode { get; set; }
        public string ApplyVAT { get; set; }
        public string TaxDate { get; set; }
        public int Series { get; set; }
        public object BankCode { get; set; }
        public object BankAccount { get; set; }
        public float DiscountPercent { get; set; }
        public object ProjectCode { get; set; }
        public string CurrencyIsLocal { get; set; }
        public float DeductionPercent { get; set; }
        public float DeductionSum { get; set; }
        public float CashSumFC { get; set; }
        public float CashSumSys { get; set; }
        public object BoeAccount { get; set; }
        public float BillOfExchangeAmount { get; set; }
        public object BillofExchangeStatus { get; set; }
        public float BillOfExchangeAmountFC { get; set; }
        public float BillOfExchangeAmountSC { get; set; }
        public object BillOfExchangeAgent { get; set; }
        public object WTCode { get; set; }
        public float WTAmount { get; set; }
        public float WTAmountFC { get; set; }
        public float WTAmountSC { get; set; }
        public object WTAccount { get; set; }
        public float WTTaxableAmount { get; set; }
        public string Proforma { get; set; }
        public object PayToBankCode { get; set; }
        public object PayToBankBranch { get; set; }
        public object PayToBankAccountNo { get; set; }
        public object PayToCode { get; set; }
        public object PayToBankCountry { get; set; }
        public string IsPayToBank { get; set; }
        public int DocEntry { get; set; }
        public string PaymentPriority { get; set; }
        public object TaxGroup { get; set; }
        public float BankChargeAmount { get; set; }
        public float BankChargeAmountInFC { get; set; }
        public float BankChargeAmountInSC { get; set; }
        public float UnderOverpaymentdifference { get; set; }
        public float UnderOverpaymentdiffSC { get; set; }
        public float WtBaseSum { get; set; }
        public float WtBaseSumFC { get; set; }
        public float WtBaseSumSC { get; set; }
        public string VatDate { get; set; }
        public string TransactionCode { get; set; }
        public string PaymentType { get; set; }
        public float TransferRealAmount { get; set; }
        public string DocObjectCode { get; set; }
        public string DocTypte { get; set; }
        public string DueDate { get; set; }
        public object LocationCode { get; set; }
        public string Cancelled { get; set; }
        public string ControlAccount { get; set; }
        public float UnderOverpaymentdiffFC { get; set; }
        public string AuthorizationStatus { get; set; }
        public object BPLID { get; set; }
        public object BPLName { get; set; }
        public object VATRegNum { get; set; }
        public object BlanketAgreement { get; set; }
        public string PaymentByWTCertif { get; set; }
        public object Cig { get; set; }
        public object Cup { get; set; }
        public float U_DtoImp { get; set; }
        public string U_Usuario { get; set; }
        public float U_DctoPrc { get; set; }
        public string U_CODIGODGI { get; set; }
        public float U_TASA { get; set; }
        public object U_IMPUESTO { get; set; }
        public string U_TIPOCFE_REFERENCIA { get; set; }
        public object U_NroFac { get; set; }
        public object U_FechaFac { get; set; }
        public object U_TIPOCFE { get; set; }
        public object U_ADJUNTO { get; set; }
        public object U_EnvRec { get; set; }
        public object U_PagoRef { get; set; }
        public object[] PaymentChecks { get; set; }
        public object[] PaymentInvoices { get; set; }
        public Paymentcreditcard[] PaymentCreditCards { get; set; }
        public Paymentaccount[] PaymentAccounts { get; set; }
        public Billofexchange BillOfExchange { get; set; }
        public object[] WithholdingTaxCertificatesCollection { get; set; }
        public object[] ElectronicProtocols { get; set; }
        public object[] CashFlowAssignments { get; set; }
        public object[] Payments_ApprovalRequests { get; set; }
        public object[] WithholdingTaxDataWTXCollection { get; set; }
    }

    public class Billofexchange
    {
    }

    public class Paymentcreditcard
    {
        public int LineNum { get; set; }
        public int CreditCard { get; set; }
        public string CreditAcct { get; set; }
        public string CreditCardNumber { get; set; }
        public string CardValidUntil { get; set; }
        public string VoucherNum { get; set; }
        public object OwnerIdNum { get; set; }
        public object OwnerPhone { get; set; }
        public int PaymentMethodCode { get; set; }
        public int NumOfPayments { get; set; }
        public string FirstPaymentDue { get; set; }
        public float FirstPaymentSum { get; set; }
        public float AdditionalPaymentSum { get; set; }
        public float CreditSum { get; set; }
        public string CreditCur { get; set; }
        public float CreditRate { get; set; }
        public object ConfirmationNum { get; set; }
        public int NumOfCreditPayments { get; set; }
        public string CreditType { get; set; }
        public string SplitPayments { get; set; }
    }

    public class Paymentaccount
    {
        public int LineNum { get; set; }
        public string AccountCode { get; set; }
        public float SumPaid { get; set; }
        public float SumPaidFC { get; set; }
        public string Decription { get; set; }
        public object VatGroup { get; set; }
        public string AccountName { get; set; }
        public float GrossAmount { get; set; }
        public object ProfitCenter { get; set; }
        public object ProjectCode { get; set; }
        public float VatAmount { get; set; }
        public object ProfitCenter2 { get; set; }
        public object ProfitCenter3 { get; set; }
        public object ProfitCenter4 { get; set; }
        public object ProfitCenter5 { get; set; }
        public object LocationCode { get; set; }
        public float EqualizationVatAmount { get; set; }
    }

}

