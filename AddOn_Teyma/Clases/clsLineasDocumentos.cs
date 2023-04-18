using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOn_Caja.Clases
{
    public class clsLineasDocumentos
    {
        public int DocEntry;
        public int DocNum;
        public int LineNum;
        public int ObjType;
        public string TaxCode;
        public string ItemCode;
        public string CardCode;
        public string CardName;
        public string Articulo;
        public decimal Total;
        public decimal TotalIVA;
        public decimal TotalConIVA;
        public double Cantidad;
        public string Moneda;
        public DateTime Fecha;
        public bool Check;
        public string ArtPropExclusivo;
        public double Descuento;
        public string ItemName;

        public int NCDocEntry;
        public int NCDocNum;
        public int NCLineNum;
    }
}
