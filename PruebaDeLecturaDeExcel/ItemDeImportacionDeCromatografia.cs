using System;

namespace PruebaDeLecturaDeExcel {
    public class ItemDeImportacionDeCromatografia {
        public string? PuntoDeEntrega { get; set; }
        public DateTime Fecha { get; set; }
        public string? CodigoDeComposicion { get; set; }
        public string? Definicion { get; set; }
        public decimal Molar { get; set; }
    }
}
