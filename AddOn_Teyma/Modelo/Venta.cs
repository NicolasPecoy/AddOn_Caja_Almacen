using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOn_Caja.Modelo
{
    class Venta
    {
        public string operacion { get; set; }
        public bool facturaConsumidorFinal { get; set; }
        public double facturaNro { get; set; }

        public double facturaMonto { get; set; }
        public double facturaMontoGravado { get; set; }
        public double facturaMontoIVA { get; set; }
        public int tarjetaId { get; set; }
        public string monedaISO { get; set; }
        public int multiEmp { get; set; }
        public Terminal terminal { get; set; }

        //decreto de ley
        public int decretoLey { get; set; }

        // Variables Credito
        public int cuotas { get; set; }

        // variables cancelacion

        public int ticketNro { get; set; }

  

        //Constructor contado
        public Venta(string operacion, bool facturaConsumidorFinal, double facturaNro, double facturaMonto, double facturaMontoGravado, double facturaMontoIVA, int tarjetaId, string monedaISO, int multiEmp, Terminal terminal, int cuota, int decretoLey)
        {
            this.operacion = operacion;
            this.facturaConsumidorFinal = facturaConsumidorFinal;
            this.facturaNro = facturaNro;
            this.facturaMonto = facturaMonto;
            this.facturaMontoGravado = facturaMontoGravado;
            this.facturaMontoIVA = facturaMontoIVA;
            this.tarjetaId = tarjetaId;
            this.monedaISO = monedaISO;
            this.multiEmp = multiEmp;
            this.terminal = terminal;
            this.cuotas = cuota;
            this.decretoLey = decretoLey;

        }

        //constructor devolucion

        public Venta(string operacion, int ticketNro, double facturaNro, double monto, string moneda, Terminal terminal, double facturaMontoGravado, double iva)
        {
            this.operacion = operacion;
            this.ticketNro = ticketNro;
            this.facturaMonto = monto;
            this.facturaNro = facturaNro;
            this.monedaISO = moneda;
            this.terminal = terminal;
            this.facturaMontoGravado = facturaMontoGravado;
            this.facturaMontoIVA = iva;

        }

        public Venta()
        {
        }

        ////Constructor Credito
        //public Venta(string operacion, bool facturaConsumidorFinal, double facturaNro, double facturaMonto, double facturaMontoGravado, double facturaMontoIVA, int tarjetaId, string monedaISO, int multiEmp, Terminal terminal,
        //            int cuotas)
        //{
        //    this.operacion = operacion;
        //    this.facturaConsumidorFinal = facturaConsumidorFinal;
        //    this.facturaNro = facturaNro;
        //    this.facturaMonto = facturaMonto;
        //    this.facturaMontoGravado = facturaMontoGravado;
        //    this.facturaMontoIVA = facturaMontoIVA;
        //    this.tarjetaId = tarjetaId;
        //    this.monedaISO = monedaISO;
        //    this.multiEmp = multiEmp;
        //    this.terminal = terminal;
        //    this.cuotas = cuotas;
        //    this.titular = titular;
        //    this.vencimiento = vencimiento;
        //    this.cvc = cvc;
        //    this.numtarjeta = numtarjeta;
        //    this.cedula = cedula;



        //}
    }
}
