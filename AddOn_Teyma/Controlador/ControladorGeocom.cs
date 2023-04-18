using AddOn_Caja.Clases;
using AddOn_Caja.Sistema;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOn_Caja.Controlador
{
    public class ControladorGeocom
    {
     
        SistemaGeocom SG = new SistemaGeocom();
        SboClass sbo;

        public ControladorGeocom(SboClass sb)
        {
            sbo = sb;
        }

        public ControladorGeocom()
        {
           
        }

        //Operacion venta

        public GeocomWSProductivo.PurchaseQueryResponse enviarVentaPosGeocom(double montoTotalPago, String PosID, String moneda, double montoGravado, double montoIVA, string nroFactura, int cuotas, int decretoLey, double montoTotalFactura, string systemId, string Branch, string clientAppId, string userId, string merchant, string plan, string sello)
        {
            GeocomWSProductivo.PurchaseQueryResponse respuestaTransaccion = SG.enviarVentaPosGeocom(montoTotalPago,PosID,moneda,montoGravado,montoIVA,nroFactura,cuotas,decretoLey,montoTotalFactura,systemId,Branch,clientAppId,userId, merchant, plan, sello);

            return respuestaTransaccion;
        }

        //Operación Cancelación

        public GeocomWSProductivo.PurchaseQueryResponse cancelacion(double montoTotalPago, String PosID, String moneda, double montoGravado, double montoIVA, string nroFactura, int cuotas, int decretoLey, double montoTotalFactura, string systemId, string Branch, string clientAppId, string userId, string ticket, LogGeocom objetoLogVenta)
        {
            GeocomWSProductivo.PurchaseQueryResponse respuestaTransaccion = SG.cancelacion(montoTotalPago, PosID, moneda, montoGravado, montoIVA, nroFactura, cuotas, decretoLey, montoTotalFactura, systemId, Branch, clientAppId, userId, ticket, objetoLogVenta);

            return respuestaTransaccion;
        }


        //Operación Devolución devolucion

        public GeocomWSProductivo.PurchaseQueryResponse devolucion(double montoTotalPago, String PosID, String moneda, double montoGravado, double montoIVA, string nroFactura, int cuotas, int decretoLey, double montoTotalFactura, string systemId, string Branch, string clientAppId, string userId, string ticket, LogGeocom objetoLogVenta)
        {
            GeocomWSProductivo.PurchaseQueryResponse respuestaTransaccion = SG.devolucion(montoTotalPago, PosID, moneda, montoGravado, montoIVA, nroFactura, cuotas, decretoLey, montoTotalFactura, systemId, Branch, clientAppId, userId, ticket, objetoLogVenta);

            return respuestaTransaccion;
        }
    }
}
