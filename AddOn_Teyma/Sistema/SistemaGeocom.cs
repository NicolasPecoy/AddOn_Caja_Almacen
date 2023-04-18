using AddOn_Caja.Clases;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOn_Caja.Sistema
{
    class SistemaGeocom
    {
        public SistemaGeocom()
        {
        }

        public GeocomWSProductivo.PurchaseQueryResponse enviarVentaPosGeocom(double montoTotalPago, String PosID, String moneda, double montoGravado, double montoIVA, string nroFactura, int cuotas, int decretoLey, double montoTotalFactura, string systemId, string Branch, string clientAppId, string userId, string merchant, string plan, string sello)
        {
            DateTime fecha = DateTime.Now;
            // Instancia de interface
            GeocomWSProductivo.IPOSInterfaceService pOSInterfaceService = new GeocomWSProductivo.POSInterfaceServiceClient();
            GeocomWSProductivo.PurchaseQueryResponse respuestaPool = null;
            GeocomWSProductivo.PurchaseQueryRequest transaccionEstado = new GeocomWSProductivo.PurchaseQueryRequest();
            //Instancia objeto transaccion
            GeocomWSProductivo.PurchaseRequest test = new GeocomWSProductivo.PurchaseRequest();

            if (moneda.Equals("UYU") || moneda.Equals("CLP") || moneda.Equals("ARS") || moneda.Equals("$"))
                moneda = "858";
            else
                moneda = "840"; //dolares
            
            if (cuotas == 0)
                cuotas = 1;

            test.PosID = PosID;
            test.SystemId = systemId;
            test.Branch = Branch;
            test.ClientAppId = clientAppId;
            test.UserId = userId;
            test.TransactionDateTimeyyyyMMddHHmmssSSS = fecha.ToString();
            
            if (!String.IsNullOrEmpty(merchant))
                test.Merchant = merchant;
           
            test.Amount = (montoTotalPago * 100).ToString();
            test.Quotas = cuotas;
            if (!String.IsNullOrEmpty(plan))
                test.Plan = Convert.ToInt32(plan); 
            else
                test.Plan = 0;
          
            test.Currency = moneda;
            test.TaxRefund = decretoLey;
            test.TaxableAmount = (Math.Round(montoGravado,2) * 100).ToString();
            test.InvoiceAmount = (Math.Round(montoTotalFactura) * 100).ToString();
            test.InvoiceNumber = nroFactura;
            //Almacen , validas Issuer
            test.Issuer = sello;
            //Time Out
            test.TransactionTimeout = 90;

            try
            {
                //se envia venta
                GeocomWSProductivo.PurchaseResponse respuesta = pOSInterfaceService.processFinancialPurchase(test);

                // estadoTransaccionCod = respuesta.ResponseCode;
                transaccionEstado.STransactionId = respuesta.STransactionId;
                //TEST
                transaccionEstado.Branch = test.Branch;
                transaccionEstado.ClientAppId = test.ClientAppId;
                transaccionEstado.ExtensionData = test.ExtensionData;
                transaccionEstado.PosID = test.PosID;
                transaccionEstado.STransactionId = respuesta.STransactionId;
                transaccionEstado.SystemId = test.SystemId;
                transaccionEstado.TransactionDateTimeyyyyMMddHHmmssSSS = test.TransactionDateTimeyyyyMMddHHmmssSSS;
                transaccionEstado.TransactionId = respuesta.TransactionId;
                transaccionEstado.UserId = test.UserId;

                //se manda a verificar estado GeocomWS.PurchaseQueryResponse 
                respuestaPool = pOSInterfaceService.processFinancialPurchaseQuery(transaccionEstado);
                //se valida estado de transaccion
                codigoRespuesta(respuestaPool.ResponseCode);

                if (respuestaPool.ResponseCode != 0 && respuestaPool.ResponseCode != 10 && respuestaPool.ResponseCode != 12)
                    return respuestaPool;
                else if (respuestaPool.ResponseCode == 0)
                    return respuestaPool;
                else
                {
                    System.Threading.Thread.Sleep(500);
                    bool salir = false;
                  
                    while (respuestaPool.ResponseCode != 0)
                    {
                        System.Threading.Thread.Sleep(2000);

                        respuestaPool = pOSInterfaceService.processFinancialPurchaseQuery(transaccionEstado);
                        if (respuestaPool.ResponseCode != 0 && respuestaPool.ResponseCode != 10 && respuestaPool.ResponseCode != 12 && respuestaPool.ResponseCode != 11)
                            salir = true;
                        else if (respuestaPool.ResponseCode == 0)
                            salir = true;
                        else if (respuestaPool.ResponseCode == 0)
                            salir = true;
                        else if (respuestaPool.ResponseCode == 11)
                            return respuestaPool;
                    }

                    return respuestaPool;
                }
            }
            catch (Exception ex)
            {
                bool retorno = ProcessFinancialReverse(PosID, systemId, Branch, clientAppId, userId, transaccionEstado.TransactionId, transaccionEstado.STransactionId);
                if (retorno)
                    respuestaPool.PosResponseCode = "-1";

                return respuestaPool;
            }
        }
        
        public GeocomWSProductivo.PurchaseQueryResponse cancelacion(double montoTotalPago, String PosID, String moneda, double montoGravado, double montoIVA, string nroFactura, int cuotas, int decretoLey, double montoTotalFactura, string systemId, string Branch, string clientAppId, string userId, string ticket, LogGeocom objetoLogVenta)
        {
            DateTime fecha = DateTime.Now;
            DateTime fechaTransaccion = DateTime.Now;
            // Instancia de interface
            GeocomWSProductivo.IPOSInterfaceService pOSInterfaceService = new GeocomWSProductivo.POSInterfaceServiceClient();
            GeocomWSProductivo.PurchaseQueryResponse respuestaPool = null;
            GeocomWSProductivo.PurchaseQueryRequest transaccionEstado = new GeocomWSProductivo.PurchaseQueryRequest();
            //Instancia objeto transaccion
            GeocomWSProductivo.PurchaseVoidRequest test = new GeocomWSProductivo.PurchaseVoidRequest();
            //200520
            try
            {
                if (!string.IsNullOrEmpty(objetoLogVenta.TransactionDateTime))
                {
                    string anio = "20" + objetoLogVenta.TransactionDateTime.Substring(0, 2);
                    string mes = objetoLogVenta.TransactionDateTime.Substring(2, objetoLogVenta.TransactionDateTime.Length - 4);
                    string dia = objetoLogVenta.TransactionDateTime.Substring(4);

                    string hora = objetoLogVenta.transactionDateTime.Substring(0, 2);
                    string minutos = objetoLogVenta.transactionDateTime.Substring(2, objetoLogVenta.transactionDateTime.Length - 4);
                    string segundos = objetoLogVenta.transactionDateTime.Substring(4);

                    fechaTransaccion = new DateTime(Convert.ToInt32(anio), Convert.ToInt32(mes), Convert.ToInt32(dia), Convert.ToInt32(hora), Convert.ToInt32(minutos), Convert.ToInt32(segundos));
                }
            }
            catch (Exception ex)
            {
               
            }

            test.PosID = PosID;
            test.SystemId = systemId;
            test.Branch = Branch;
            test.ClientAppId = clientAppId;
            test.UserId = userId;
            test.TransactionDateTimeyyyyMMddHHmmssSSS = fecha.ToString();
            test.TicketNumber = ticket;  //objetoLogVenta.ticket.ToString();

            //Time Out
            test.TransactionTimeout = 90;

            try
            {
                //se envia cancelacion
                GeocomWSProductivo.PurchaseVoidResponse respuesta = pOSInterfaceService.processFinancialPurchaseVoidByTicket(test);

                // estadoTransaccionCod = respuesta.ResponseCode;
                transaccionEstado.STransactionId = respuesta.STransactionId;
                //TEST
                transaccionEstado.Branch = test.Branch;
                transaccionEstado.ClientAppId = test.ClientAppId;
                transaccionEstado.ExtensionData = test.ExtensionData;
                transaccionEstado.PosID = test.PosID;
                transaccionEstado.STransactionId = respuesta.STransactionId;
                transaccionEstado.SystemId = test.SystemId;
                transaccionEstado.TransactionDateTimeyyyyMMddHHmmssSSS = test.TransactionDateTimeyyyyMMddHHmmssSSS;
                transaccionEstado.TransactionId = respuesta.TransactionId;
                transaccionEstado.UserId = test.UserId;


                //se manda a verificar estado GeocomWS.PurchaseQueryResponse 
                respuestaPool = pOSInterfaceService.processFinancialPurchaseQuery(transaccionEstado);
                //se valida estado de transaccion
                codigoRespuesta(respuestaPool.ResponseCode);


                if (respuestaPool.ResponseCode != 0 && respuestaPool.ResponseCode != 10 && respuestaPool.ResponseCode != 12)
                {

                    return respuestaPool;

                }
                else if (respuestaPool.ResponseCode == 0)
                {
                    return respuestaPool;

                }
                else
                {
                    System.Threading.Thread.Sleep(500);
                    bool salir = false;
                 
                    while (respuestaPool.ResponseCode != 0)
                    {
                        System.Threading.Thread.Sleep(2000);
                        respuestaPool = pOSInterfaceService.processFinancialPurchaseQuery(transaccionEstado);
                        if (respuestaPool.ResponseCode != 0 && respuestaPool.ResponseCode != 10 && respuestaPool.ResponseCode != 12 && respuestaPool.ResponseCode != 11)
                        {

                            salir = true;

                        }
                        else if (respuestaPool.ResponseCode == 0)
                        {
                            salir = true;
                        }
                        else if (respuestaPool.ResponseCode == 11)
                        {
                            return respuestaPool;
                        }
                     
                 

                    }

                    return respuestaPool;
                }



            }
            catch (Exception ex)
            {
                bool retorno = ProcessFinancialReverse(PosID, systemId, Branch, clientAppId, userId, transaccionEstado.TransactionId, transaccionEstado.STransactionId);
                if (retorno)
                {
                    respuestaPool.PosResponseCode = "-1";

                }

                return respuestaPool;
            }


        }

        //este metodo se dispara cuando hay una excepcion es SAP o Geocom
        public bool ProcessFinancialReverse(String PosID, string systemId, string Branch, string clientAppId, string userId, long TransactionId, string STransactionId)
        {
            bool retorno = false;
            DateTime fecha = DateTime.Now;
            // Instancia de interface
            GeocomWSProductivo.IPOSInterfaceService pOSInterfaceService = new GeocomWSProductivo.POSInterfaceServiceClient();
            //Instancia objeto transaccion
            GeocomWSProductivo.ReverseRequest reversa = new GeocomWSProductivo.ReverseRequest();


            reversa.PosID = PosID;
            reversa.SystemId = systemId;
            reversa.Branch = Branch;
            reversa.ClientAppId = clientAppId;
            reversa.UserId = userId;
            reversa.TransactionDateTimeyyyyMMddHHmmssSSS = fecha.ToString();
            reversa.TransactionId = TransactionId;
            reversa.STransactionId = STransactionId;


            try
            {
                //se envia cancelacion
                GeocomWSProductivo.ReverseResponse respuesta = pOSInterfaceService.processFinancialReverse(reversa);

                codigoRespuesta(respuesta.ResponseCode);
                if (respuesta.ResponseCode == 0)
                {
                    retorno = true;
                }


            }
            catch (Exception ex)
            {

                //  return respuestaPool;
            }

            return retorno;
        }

        public GeocomWSProductivo.PurchaseQueryResponse devolucion(double montoTotalPago, String PosID, String moneda, double montoGravado, double montoIVA, string nroFactura, int cuotas, int decretoLey, double montoTotalFactura, string systemId, string Branch, string clientAppId, string userId, string ticket, LogGeocom objetoLogVenta)
        {

            DateTime fecha = DateTime.Now;
            // Instancia de interface
            GeocomWSProductivo.IPOSInterfaceService pOSInterfaceService = new GeocomWSProductivo.POSInterfaceServiceClient();
            GeocomWSProductivo.PurchaseQueryResponse respuestaPool = null;

            //Instancia objeto transaccion
            GeocomWSProductivo.PurchaseRefundRequest test = new GeocomWSProductivo.PurchaseRefundRequest();
            //200520
            try
            {
                string anio = "20" + objetoLogVenta.TransactionDateTime.Substring(0, 2);
                string mes = objetoLogVenta.TransactionDateTime.Substring(2, objetoLogVenta.TransactionDateTime.Length - 4);
                string dia = objetoLogVenta.TransactionDateTime.Substring(4);

                string hora = objetoLogVenta.transactionDateTime.Substring(0, 2);
                string minutos = objetoLogVenta.transactionDateTime.Substring(2, objetoLogVenta.transactionDateTime.Length - 4);
                string segundos = objetoLogVenta.transactionDateTime.Substring(4);

                DateTime fechaTransaccion = new DateTime(Convert.ToInt32(anio), Convert.ToInt32(mes), Convert.ToInt32(dia), Convert.ToInt32(hora), Convert.ToInt32(minutos), Convert.ToInt32(segundos));
                test.OriginalTransactionDateyyMMdd = objetoLogVenta.TransactionDateTime.Substring(0, 2) + mes + dia; //Obligatorio
            }
            catch (Exception)
            {
                
               
            }
          

            test.PosID = PosID; //Obligatorio
            test.SystemId = systemId; //Obligatorio
            test.Branch = Branch; //Obligatorio
            test.ClientAppId = clientAppId; //Obligatorio
            test.UserId = userId; //Obligatorio
            test.TransactionDateTimeyyyyMMddHHmmssSSS = fecha.ToString(); //Obligatorio
            test.TicketNumber = objetoLogVenta.ticket.ToString(); //Obligatorio
           
            test.Amount = (montoTotalPago * 100).ToString(); //Obligatorio
            test.Quotas = cuotas; //Obligatorio

            if (!String.IsNullOrEmpty(objetoLogVenta.Merchant))
            {
                test.Merchant = objetoLogVenta.Merchant;
            }
            else if (String.IsNullOrEmpty(objetoLogVenta.Merchant))
            {
                test.Merchant = "";
            }

            test.Amount = (montoTotalPago * 100).ToString();
            test.Quotas = cuotas;
            if (!String.IsNullOrEmpty(objetoLogVenta.plan))
            {
                test.Plan = Convert.ToInt32(objetoLogVenta.plan);
            }
            else
            {
                test.Plan = 0;
            }
        
            test.Currency = objetoLogVenta.monedaTransaccionCod; //Obligatorio
            test.TaxRefund = decretoLey;
            test.TaxableAmount = (montoGravado * 100).ToString(); //Obligatorio
            test.InvoiceAmount = (montoTotalFactura * 100).ToString(); //Obligatorio
            test.InvoiceNumber = nroFactura; //Obligatorio


            //Time Out
            test.TransactionTimeout = 90;

            try
            {
                //se envia venta
                GeocomWSProductivo.PurchaseRefundResponse respuesta = pOSInterfaceService.processFinancialPurchaseRefund(test);


                // estadoTransaccionCod = respuesta.ResponseCode;

                GeocomWSProductivo.PurchaseQueryRequest transaccionEstado = new GeocomWSProductivo.PurchaseQueryRequest();
                transaccionEstado.STransactionId = respuesta.STransactionId;
                //TEST
                transaccionEstado.Branch = test.Branch;
                transaccionEstado.ClientAppId = test.ClientAppId;
                transaccionEstado.ExtensionData = test.ExtensionData;
                transaccionEstado.PosID = test.PosID;
                transaccionEstado.STransactionId = respuesta.STransactionId;
                transaccionEstado.SystemId = test.SystemId;
                transaccionEstado.TransactionDateTimeyyyyMMddHHmmssSSS = test.TransactionDateTimeyyyyMMddHHmmssSSS;
                transaccionEstado.TransactionId = respuesta.TransactionId;
                transaccionEstado.UserId = test.UserId;




                //se manda a verificar estado GeocomWS.PurchaseQueryResponse 
                respuestaPool = pOSInterfaceService.processFinancialPurchaseQuery(transaccionEstado);
                //se valida estado de transaccion
                codigoRespuesta(respuestaPool.ResponseCode);


                if (respuestaPool.ResponseCode != 0 && respuestaPool.ResponseCode != 10 && respuestaPool.ResponseCode != 12)
                {

                    return respuestaPool;
                    Console.WriteLine("Error en transaccion");
                    Console.ReadLine();

                }
                else if (respuestaPool.ResponseCode == 0)
                {
                    return respuestaPool;
                    Console.WriteLine("Transaccion Exitosa");
                    Console.ReadLine();
                }
                else
                {
                    System.Threading.Thread.Sleep(500);
                    bool salir = false;

                    while (respuestaPool.ResponseCode != 0)
                    {
                        respuestaPool = pOSInterfaceService.processFinancialPurchaseQuery(transaccionEstado);
                        if (respuestaPool.ResponseCode != 0 && respuestaPool.ResponseCode != 10 && respuestaPool.ResponseCode != 12 && respuestaPool.ResponseCode != 11)
                        {
                            System.Threading.Thread.Sleep(2000);


                            salir = true;

                        }
                        else if (respuestaPool.ResponseCode == 0)
                        {
                            salir = true;
                        }
                        else if (respuestaPool.ResponseCode == 11)
                        {
                            return respuestaPool;
                        }


                    }

                    return respuestaPool;
                }



            }
            catch (Exception ex)
            {

                //  return respuestaPool;
            }

            return respuestaPool;
        }

        static public void codigoRespuesta(int codigo)
        {
            switch (codigo)
            {
                case 0:
                    Console.WriteLine("Resultado OK");
                    break;
                case 100:
                    Console.WriteLine("Número de pinpad inválido");
                    break;
                case 101:
                    Console.WriteLine("Número de sucursal inválido");
                    break;
                case 102:
                    Console.WriteLine("Número de caja inválido");
                    break;
                case 103:
                    Console.WriteLine("Fecha de la transacción inválida");
                    break;
                case 104:
                    Console.WriteLine("Monto no válido");
                    break;
                case 105:
                    Console.WriteLine("Cantidad de cuotas inválidas");
                    break;
                case 106:
                    Console.WriteLine("Número de plan inválido");
                    break;
                case 107:
                    Console.WriteLine("Número de factura inválido");
                    break;
                case 108:
                    Console.WriteLine("Moneda ingresada no válida");
                    break;
                case 109:
                    Console.WriteLine("Número de ticket inválido");
                    break;
                case 110:
                    Console.WriteLine("No existe transacción.");
                    break;
                case 111:
                    Console.WriteLine("Transacción finalizada.");
                    break;
                case 112:
                    Console.WriteLine("Identificador de sistema inválido.");
                    break;
                case 113:
                    Console.WriteLine("Se debe consultar por la transacción");
                    break;
                case 10:
                    Console.WriteLine("Aguardando por operación en el pinpad.");
                    break;
                case 11:
                    Console.WriteLine("Tiempo de transacción excedido, envíe datos nuevamente");
                    break;
                case 12:
                    Console.WriteLine("Pinpad consultó datos (se pasó la tarjeta).");
                    break;
                case 999:
                    Console.WriteLine("Error no determinado.");
                    break;

            }
            Console.ReadLine();
        }

    }
}
