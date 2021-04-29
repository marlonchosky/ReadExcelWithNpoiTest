using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace PruebaDeLecturaDeExcel {
    public partial class MainWindow : Window {
        public MainWindow() => this.InitializeComponent();

        private async void Button_Click(object sender, RoutedEventArgs e) {
            var rutaDelArchivo = Path.Combine(Directory.GetCurrentDirectory(), "Qualidade por Associação - 022021 - REV0 - copia.xlsx");

            var sw = Stopwatch.StartNew();
            var data = this.DeserealizarArchivo(rutaDelArchivo);
            sw.Stop();

            // Simular una tarea.
            await Task.Delay(200);

            MessageBox.Show($"Cantidad: {data.Count}, y tiempo: {sw.ElapsedMilliseconds} ms");
        }

        private IReadOnlyList<ItemDeImportacionDeCromatografia> DeserealizarArchivo(string rutaDelArchivo) {
            
            using var stream = new FileStream(rutaDelArchivo, FileMode.Open, FileAccess.Read) {
                Position = 0
            };
            var xssWorkbook = new XSSFWorkbook(stream);
            var hojaDeTrabajo = xssWorkbook.GetSheetAt(0);
            var listaDeItems = new List<ItemDeImportacionDeCromatografia>();
            for (var i = hojaDeTrabajo.FirstRowNum + 1; i <= hojaDeTrabajo.LastRowNum; i++) {
                this.AgregarItem(hojaDeTrabajo, i, listaDeItems);
            }

            return listaDeItems;
        }

        private void AgregarItem(ISheet hojaDeTrabajo, int numeroDeFila, IList<ItemDeImportacionDeCromatografia> listaDeItems) {
            var fila = hojaDeTrabajo.GetRow(numeroDeFila);
            if (fila is null) return;

            var celdaSamplingPoint = fila.GetCell(0);
            var celdaFecha = fila.GetCell(1);
            var celdaCodigoDeComposicion = fila.GetCell(2);
            var celdaDefinicion = fila.GetCell(3);
            var celdaMolar = fila.GetCell(4);

            if (AlgunaCeldaEsVacia(celdaSamplingPoint, celdaFecha, celdaCodigoDeComposicion,
                celdaDefinicion, celdaMolar)) return;

            var item = new ItemDeImportacionDeCromatografia {
                PuntoDeEntrega = celdaSamplingPoint.StringCellValue,
                Fecha = ObtenerFecha(celdaFecha),
                CodigoDeComposicion = celdaCodigoDeComposicion.StringCellValue,
                Definicion = celdaDefinicion.StringCellValue,
                Molar = (decimal)celdaMolar.NumericCellValue,
            };
            listaDeItems.Add(item);
        }

        private static DateTime ObtenerFecha(ICell celda) {
            try {
                return celda.DateCellValue;
            } catch (InvalidOperationException) {
                var textoDeCelda = celda.StringCellValue;
                return DateTime.ParseExact(textoDeCelda, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
        }

        private static bool AlgunaCeldaEsVacia(params ICell?[] celdas) => celdas.Any(c => LaCeldaEsVacia(c));

        private static bool LaCeldaEsVacia(ICell? celda) => 
            celda is null || celda.CellType == CellType.Blank || string.IsNullOrWhiteSpace(celda.ToString());
    }
}
