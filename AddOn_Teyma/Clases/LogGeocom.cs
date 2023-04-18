using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOn_Caja.Clases
{
    public class LogGeocom
    {
        public LogGeocom()
        {
            this.terminal = "-";
            this.codigoAutorizacion = "-";
            this.lote = "-";
            this.numerotarjeta = "-";
            this.nombre = "-";
            this.ci = "-";
            this.monedaTransaccionCod = "-";
            this.monedaTransaccionDescrip = "-";
            this.cuentaTarjeta = "-";
            this.nombreTarjeta = "-";
            this.selloCod = "-";
            this.selloDescripcion = "-";
            this.issuerCode = "-";
            this.issuerCodeDescripcion = "-";
            this.cardtype = "-";
            this.plan = "-";
            this.posId = "-";
            this.codigoRespuestaPos = "-";
            this.codigoRespuestaPosDescripcion = "-";
            this.cuotas = "-";
            this.impuestocodigo = "-";
            this.ticket = "";
            this.monto = "-";
            this.transaccionType = "-";
            this.transaccionTypeDescripcion = "-";
            this.codigoTarjetaSAP = "-";
            this.EstatusGeocomTransaccion = "";
            this.EstatusSAPTransaccion = "";
            this.Merchant = "";
        }

        public string terminal { get; set; }
        public string codigoAutorizacion { get; set; }
        public string lote { get; set; }
        public string numerotarjeta { get; set; }
        public string nombre { get; set; }
        public string ci { get; set; }
        public string monedaTransaccionCod { get; set; }
        public string monedaTransaccionDescrip { get; set; }
        public string cuentaTarjeta { get; set; }
        public string nombreTarjeta { get; set; }
        public string selloCod { get; set; }
        public string selloDescripcion { get; set; }
        public string issuerCode { get; set; }
        public string issuerCodeDescripcion { get; set; }
        public string cardtype { get; set; }
        public string plan { get; set; }
        public string posId { get; set; }
        public string codigoRespuestaPos { get; set; }
        public string codigoRespuestaPosDescripcion { get; set; }
        public string cuotas { get; set; }
        public string impuestocodigo { get; set; }
        public string ticket { get; set; }
        public string monto { get; set; }

        public string TransactionDateTime { get; set; }
        public DateTime fechaTransaccion { get; set; }
        public string transaccionType { get; set; }
        public string transaccionTypeDescripcion { get; set; }
        public string codigoTarjetaSAP { get; set; }
        public string transactionDateTime { get; set; }
        public string TaxableAmount { get; set; }
        public string TaxRefund { get; set; }
        public string InvoiceAmount { get; set; }
        public string EstatusGeocomTransaccion { get; set; }
        public string EstatusSAPTransaccion { get; set; }
        public string Merchant { get; set; }


    }
}
