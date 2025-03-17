using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CasosDePruebaECF
{
    public static class Services
    {
        public static DataTable dt;
        public static DataRow dr;
        public static Dictionary<int, string> CasoDePrueba = new Dictionary<int, string>();

        public static string loadJson(int index)
        {
            dr = dt.Rows[index];
            return ConvertToJson();
        }

        static string ConvertToJson()
        {

            string ECF = $@"{{
    ""ECF"" : {{
        ""Encabezado"": {{
            ""Version"": ""{dr["Version"]}"",
            ""IdDoc"": {{
                ""TipoeCF"": ""{dr["TipoeCF"]}"",
                ""eNCF"": ""{dr["ENCF"]}"",
                {validarCampo("FechaVencimientoSecuencia")}
                {validarCampo("IndicadorNotaCredito")}
                {validarCampo("IndicadorEnvioDiferido")}
                {validarCampo("IndicadorMontoGravado")}
                {validarCampo("TipoIngresos")}
                {validarCampo("TipoPago")}
                {validarCampo("FechaLimitePago")}
                {validarCampo("TerminoPago")}
                {TablaFormasPago()}
                {validarCampo("TipoCuentaPago")}
                {validarCampo("NumeroCuentaPago")}
                {validarCampo("BancoPago")}
                {validarCampo("FechaDesde")}
                {validarCampo("FechaHasta")}
                {validarCampo("TotalPaginas")}
            }},
            ""Emisor"":{{
                ""RNCEmisor"": ""{dr["RNCEmisor"]}"",
                ""RazonSocialEmisor"": ""{dr["RazonSocialEmisor"]}"",
                {validarCampo("NombreComercial")}
                {validarCampo("Sucursal")}
                {validarCampo("DireccionEmisor")}
                {validarCampo("Municipio")}
                {validarCampo("Provincia")}
                {TablaTelefonoEmisor()}
                {validarCampo("CorreoEmisor")}
                {validarCampo("WebSite")}
                {validarCampo("ActividadEconomica")}
                {validarCampo("CodigoVendedor")}
                {validarCampo("NumeroFacturaInterna")}
                {validarCampo("NumeroPedidoInterno")}
                {validarCampo("ZonaVenta")}
                {validarCampo("RutaVenta")}
                {validarCampo("InformacionAdicionalEmisor")}
                {validarCampo("FechaEmision")}
            }},
            {Comprador()}
            {InformacionesAdicionales()}
            {Transporte()}
            ""Totales"":{{
                {validarCampo("MontoGravadoTotal")}
                {validarCampo("MontoGravadoI1")}
                {validarCampo("MontoGravadoI2")}
                {validarCampo("MontoGravadoI3")}
                {validarCampo("MontoExento")}
                {validarCampo("ITBIS1")}
                {validarCampo("ITBIS2")}
                {validarCampo("ITBIS3")}
                {validarCampo("TotalITBIS")}
                {validarCampo("TotalITBIS1")}
                {validarCampo("TotalITBIS2")}
                {validarCampo("TotalITBIS3")}
                {validarCampo("MontoImpuestoAdicional")}
                {ImpuestosAdicionales()}
                {validarCampo("MontoTotal")}
                {validarCampo("MontoNoFacturable")}
                {validarCampo("MontoPeriodo")}
                {validarCampo("SaldoAnterior")}
                {validarCampo("MontoAvancePago")}
                {validarCampo("ValorPagar")}
                {validarCampo("TotalITBISRetenido")}
                {validarCampo("TotalISRRetencion")}
                {validarCampo("TotalITBISPercepcion")}
                {validarCampo("TotalISRPercepcion")}
            }},
            {OtraMoneda()}
        }},
        ""DetallesItems"": {{
            {Item()}
        }},
        {DescuentosORecargos()}
        {InformacionReferencia()}
        ""FechaHoraFirma"": ""{(Convert.ToDateTime(DateTime.Now).ToString("dd-MM-yyyy HH:mm:ss", new CultureInfo("es-DO")))}""
    }}
}}";
            ECF = Regex.Replace(ECF, @",\s*}", "}");
            return string.Join("\n", ECF.Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)));
        }
        private static string DescuentosORecargos()
        {
            if (dr[$"NumeroLineaDoR[1]"].ToString() == "#e") return "";
            string DescuentosORecargos = $@"
            ""DescuentosORecargos"": {{
                ""DescuentoORecargo"" : [";

            for (int i = 1; i <= 2; i++)
            {
                if (dr[$"NumeroLineaDoR[{i}]"].ToString() != "#e")
                {
                DescuentosORecargos += $@"
                {{
                    ""NumeroLinea"": ""{dr[$"NumeroLineaDoR[{i}]"].ToString()}"",
                    {validarCampo($"TipoAjuste[{i}]", "TipoAjuste")}
                    {validarCampo($"IndicadorNorma1007[{i}]", "IndicadorNorma1007")}
                    {validarCampo($"DescripcionDescuentooRecargo[{i}]", "DescripcionDescuentooRecargo")}
                    {validarCampo($"TipoValor[{i}]", "TipoValor")}
                    {validarCampo($"ValorDescuentooRecargo[{i}]", "ValorDescuentooRecargo")}
                    {validarCampo($"MontoDescuentooRecargo[{i}]", "MontoDescuentooRecargo")}
                    {validarCampo($"MontoDescuentooRecargoOtraMoneda[{i}]", "MontoDescuentooRecargoOtraMoneda")}
                    {validarCampo($"IndicadorFacturacionDescuentooRecargo[{i}]", "IndicadorFacturacionDescuentooRecargo")}
                }},";
                }
            }

            DescuentosORecargos += $@"]
                    }},";
            return DescuentosORecargos.Replace("},]", "}]");
        }
        private static string InformacionReferencia()
        {
            if (dr[$"NCFModificado"].ToString() == "#e") return "";
            string InformacionReferencia = $@"
        ""InformacionReferencia"":{{
            ""NCFModificado"": ""{dr[$"NCFModificado"].ToString()}"",
            {validarCampo("RNCOtroContribuyente")}
            ""FechaNCFModificado"": ""{dr[$"FechaNCFModificado"].ToString()}"",
            ""CodigoModificacion"": ""{dr[$"CodigoModificacion"].ToString()}"",
            {validarCampo("RazonModificacion")}
        }},";

            return InformacionReferencia;
        }
        private static string Item()
        {
            string Item = $@"
            ""Item"" : [";

            for (int i = 1; i <= 62; i++)
            {
                if (dr[$"NumeroLinea[{i}]"].ToString() != "#e")
                {
                    Item += $@"
            {{
                ""NumeroLinea"": {dr[$"NumeroLinea[{i}]"].ToString()},
                {TablaCodigosItem(i)}
                ""IndicadorFacturacion"": {dr[$"IndicadorFacturacion[{i}]"].ToString()},
                {Retencion(i)}
                ""NombreItem"": ""{dr[$"NombreItem[{i}]"].ToString()}"",
                ""IndicadorBienoServicio"": {dr[$"IndicadorBienoServicio[{i}]"].ToString()},
                {validarCampo($"DescripcionItem[{i}]", "DescripcionItem")}
                ""CantidadItem"": ""{dr[$"CantidadItem[{i}]"].ToString()}"",
                {validarCampo($"UnidadMedida[{i}]", "UnidadMedida")}
                {validarCampo($"CantidadReferencia[{i}]", "CantidadReferencia")}
                {validarCampo($"UnidadReferencia[{i}]", "UnidadReferencia")}
                {TablaSubcantidad(i)}
                {validarCampo($"GradosAlcohol[{i}]", "GradosAlcohol")}
                {validarCampo($"PrecioUnitarioReferencia[{i}]", "PrecioUnitarioReferencia")}
                {validarCampo($"FechaElaboracion[{i}]", "FechaElaboracion")}
                {validarCampo($"FechaVencimientoItem[{i}]", "FechaVencimientoItem")}
                {Mineria(i)}
                ""PrecioUnitarioItem"": ""{dr[$"PrecioUnitarioItem[{i}]"].ToString()}"",
                {TablaSubDescuento(i)}
                {TablaSubRecargo(i)}
                {TablaImpuestoAdicional(i)}
                {OtraMonedaDetalle(i)}
                ""MontoItem"": ""{dr[$"MontoItem[{i}]"].ToString()}""
            }},";
                }
            }

            Item += $@"]";
            return Item.Replace("},]", "}]").Replace(@"""""", @"\""""");

        }
        private static string OtraMonedaDetalle(int linea)
        {
            if (dr[$"PrecioOtraMoneda[{linea}]"].ToString() == "#e") return "";
            string OtraMonedaDetalle = $@"
                    ""OtraMonedaDetalle"":{{
                        ""PrecioOtraMoneda"": ""{dr[$"PrecioOtraMoneda[{linea}]"].ToString()}"",
                        {validarCampo($"DescuentoOtraMoneda[{linea}]", "DescuentoOtraMoneda")}
                        {validarCampo($"RecargoOtraMoneda[{linea}]", "RecargoOtraMoneda")}
                        ""MontoItemOtraMoneda"": ""{dr[$"MontoItemOtraMoneda[{linea}]"].ToString()}""
                    }},";

            return OtraMonedaDetalle;
        }
        private static string TablaImpuestoAdicional(int linea)
        {
            if (dr[$"TipoImpuesto[{linea}][1]"].ToString() == "#e") return "";
            string TablaImpuestoAdicional = $@"
                    ""TablaImpuestoAdicional"": {{
                        ""ImpuestoAdicional"" : [";

            for (int i = 1; i <= 2; i++)
            {
                if (dr[$"TipoImpuesto[{linea}][{i}]"].ToString() != "#e")
                {
                    TablaImpuestoAdicional += $@"
                        {{
                            ""TipoImpuesto"": ""{dr[$"TipoImpuesto[{linea}][{i}]"].ToString()}""
                        }},";
                }
            }

            TablaImpuestoAdicional += $@"]
                    }},";
            return TablaImpuestoAdicional.Replace("},]", "}]");
        }
        private static string TablaSubRecargo(int linea)
        {
            if (dr[$"RecargoMonto[{linea}]"].ToString() == "#e") return "";
            string TablaSubRecargo = $@"
                    ""RecargoMonto"": ""{dr[$"RecargoMonto[{linea}]"].ToString()}"",
                    ""TablaSubRecargo"": {{
                        ""SubRecargo"" : [";

            for (int i = 1; i <= 5; i++)
            {
                if (dr[$"TipoSubRecargo[{linea}][{i}]"].ToString() != "#e")
                {
                    TablaSubRecargo += $@"
                        {{
                            ""TipoSubRecargo"": ""{dr[$"TipoSubRecargo[{linea}][{i}]"].ToString()}"",	
                            {validarCampo($"SubRecargoPorcentaje[{linea}][{i}]", "SubRecargoPorcentaje")}    
                            ""MontoSubRecargo"": ""{dr[$"MontoSubRecargo[{linea}][{i}]"].ToString()}""
                        }},";
                }
            }

            TablaSubRecargo += $@"]
                    }},";
            return TablaSubRecargo.Replace("},]", "}]");
        }
        private static string TablaSubDescuento(int linea)
        {
            if (dr[$"DescuentoMonto[{linea}]"].ToString() == "#e") return "";
            string TablaSubDescuento = $@"
                    ""DescuentoMonto"": ""{dr[$"DescuentoMonto[{linea}]"].ToString()}"",
                    ""TablaSubDescuento"": {{
                        ""SubDescuento"" : [";

            for (int i = 1; i <= 5; i++)
            {
                if (dr[$"TipoSubDescuento[{linea}][{i}]"].ToString() != "#e")
                {
                    TablaSubDescuento += $@"
                        {{
                            ""TipoSubDescuento"": ""{dr[$"TipoSubDescuento[{linea}][{i}]"].ToString()}"",	
                            {validarCampo($"SubDescuentoPorcentaje[{linea}][{i}]", "SubDescuentoPorcentaje")}    
                            ""MontoSubDescuento"": ""{dr[$"MontoSubDescuento[{linea}][{i}]"].ToString()}""
                        }},";
                }
            }

            TablaSubDescuento += $@"]
                    }},";
            return TablaSubDescuento.Replace("},]", "}]");
        }
        private static string Mineria(int linea)
        {
            if (dr[$"PesoNetoKilogramo[{linea}]"].ToString() == "#e") return "";
            string Mineria = $@"
                    ""Mineria"":{{
                        ""PesoNetoKilogramo"": ""{dr[$"PesoNetoKilogramo[{linea}]"].ToString()}"",
                        ""PesoNetoMineria"": ""{dr[$"PesoNetoMineria[{linea}]"].ToString()}"",
                        ""TipoAfiliacion"": ""{dr[$"TipoAfiliacion[{linea}]"].ToString()}"",
                        ""Liquidacion"": ""{dr[$"Liquidacion[{linea}]"].ToString()}""
                    }},";

            return Mineria;
        }
        private static string TablaSubcantidad(int linea)
        {
            if (dr[$"Subcantidad[{linea}][1]"].ToString() == "#e") return "";
            string TablaSubcantidad = $@"
                    ""TablaSubcantidad"": {{
                        ""SubcantidadItem"" : [";

            for (int i = 1; i <= 5; i++)
            {
                if (dr[$"Subcantidad[{linea}][{i}]"].ToString() != "#e")
                {
                    TablaSubcantidad += $@"
                        {{
                            ""Subcantidad"": ""{dr[$"Subcantidad[{linea}][{i}]"].ToString()}"",	
                            ""CodigoSubcantidad"": ""{dr[$"CodigoSubcantidad[{linea}][{i}]"].ToString()}""
                        }},";
                }
            }

            TablaSubcantidad += $@"]
                    }},";
            return TablaSubcantidad.Replace("},]", "}]");
        }
        private static string Retencion(int linea)
        {
            if (dr[$"IndicadorAgenteRetencionoPercepcion[{linea}]"].ToString() == "#e") return "";
            string Retencion = $@"
                    ""Retencion"":{{
                        ""IndicadorAgenteRetencionoPercepcion"": {dr[$"IndicadorAgenteRetencionoPercepcion[{linea}]"].ToString()},
                        {validarCampo($"MontoITBISRetenido[{linea}]", "MontoITBISRetenido")}
                        {validarCampo($"MontoISRRetenido[{linea}]", "MontoISRRetenido")}
                    }},";

            return Retencion;
        }
        private static string TablaCodigosItem(int linea)
        {
            if (dr[$"TipoCodigo[{linea}][1]"].ToString() == "#e") return "";
            string TablaCodigosItem = $@"
                    ""TablaCodigosItem"": {{
                        ""CodigosItem"" : [";

            for (int i = 1; i <= 5; i++)
            {
                if (dr[$"TipoCodigo[{linea}][{i}]"].ToString() != "#e")
                {
                    TablaCodigosItem += $@"
                        {{
                            ""TipoCodigo"": ""{dr[$"TipoCodigo[{linea}][{i}]"].ToString()}"",	
                            ""CodigoItem"": ""{dr[$"CodigoItem[{linea}][{i}]"].ToString()}""
                        }},";
                }
            }

            TablaCodigosItem += $@"]
                    }},";
            return TablaCodigosItem.Replace("},]", "}]");
        }
        private static string OtraMoneda()
        {
            if (dr["TipoMoneda"].ToString() == "#e") return "";
            string OtraMoneda = $@"
                ""OtraMoneda"":{{
                    ""TipoMoneda"": ""{dr["TipoMoneda"]}"",
                    ""TipoCambio"": ""{dr["TipoCambio"]}"",
                    {validarCampo("MontoGravadoTotalOtraMoneda")}
                    {validarCampo("MontoGravado1OtraMoneda")}
                    {validarCampo("MontoGravado2OtraMoneda")}
                    {validarCampo("MontoGravado3OtraMoneda")}
                    {validarCampo("MontoExentoOtraMoneda")}
                    {validarCampo("TotalITBISOtraMoneda")}
                    {validarCampo("TotalITBIS1OtraMoneda")}
                    {validarCampo("TotalITBIS2OtraMoneda")}
                    {validarCampo("TotalITBIS3OtraMoneda")}
                    {validarCampo("MontoImpuestoAdicionalOtraMoneda")}
                    {ImpuestosAdicionalesOtraMoneda()}
                    {validarCampo("MontoTotalOtraMoneda")}
                }},";

            return OtraMoneda;
        }
        private static string ImpuestosAdicionalesOtraMoneda()
        {
            if (dr["MontoImpuestoAdicionalOtraMoneda"].ToString() == "#e") return "";
            string ImpuestoAdicionalOtraMoneda = $@"
                ""ImpuestosAdicionalesOtraMoneda"": {{
                    ""ImpuestoAdicionalOtraMoneda"" : [";

            for (int i = 1; i <= 4; i++)
            {
                if (dr[$"TipoImpuestoOtraMoneda[{i}]"].ToString() != "#e")
                {
                    ImpuestoAdicionalOtraMoneda += $@"
                    {{
                        ""TipoImpuestoOtraMoneda"": ""{dr[$"TipoImpuestoOtraMoneda[{i}]"].ToString()}"",	
                        {validarCampo($"TasaImpuestoAdicionalOtraMoneda[{i}]", "TasaImpuestoAdicionalOtraMoneda")}
                        {validarCampo($"MontoImpuestoSelectivoConsumoEspecificoOtraMoneda[{i}]", "MontoImpuestoSelectivoConsumoEspecificoOtraMoneda")}
                        {validarCampo($"MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda[{i}]", "MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda")}
                        {validarCampo($"OtrosImpuestosAdicionalesOtraMoneda[{i}]", "OtrosImpuestosAdicionalesOtraMoneda")}
                    }},";
                }
            }

            ImpuestoAdicionalOtraMoneda += $@"]
                }},";
            return ImpuestoAdicionalOtraMoneda.Replace("},]", "}]");
        }
        private static string ImpuestosAdicionales()
        {
            if (dr["MontoImpuestoAdicional"].ToString() == "#e") return "";
            string ImpuestoAdicional = $@"
                ""ImpuestosAdicionales"": {{
                    ""ImpuestoAdicional"" : [";

            for (int i = 1; i <= 4; i++)
            {
                if (dr[$"TipoImpuesto[{i}]"].ToString() != "#e")
                {
                    ImpuestoAdicional += $@"
                    {{
                        ""TipoImpuesto"": ""{dr[$"TipoImpuesto[{i}]"].ToString()}"",	
                        {validarCampo($"TasaImpuestoAdicional[{i}]", "TasaImpuestoAdicional")}
                        {validarCampo($"MontoImpuestoSelectivoConsumoEspecifico[{i}]", "MontoImpuestoSelectivoConsumoEspecifico")}
                        {validarCampo($"MontoImpuestoSelectivoConsumoAdvalorem[{i}]", "MontoImpuestoSelectivoConsumoAdvalorem")}
                        {validarCampo($"OtrosImpuestosAdicionales[{i}]", "OtrosImpuestosAdicionales")}
                    }},";
                }
            }

            ImpuestoAdicional += $@"]
                }},";
            return ImpuestoAdicional.Replace("},]", "}]");
        }
        private static string Transporte()
        {
            for (int i = 92; i <= 105; i++)
            {
                if (dr[i].ToString() == "#e")
                {
                    if (i == 105) return "";
                }
                else
                {
                    break;
                }
            }
            string Transporte = $@"
            ""Transporte"":{{
                {validarCampo("ViaTransporte")}
                {validarCampo("PaisOrigen")}
                {validarCampo("DireccionDestino")}
                {validarCampo("PaisDestino")}
                {validarCampo("RNCIdentificacionCompaniaTransportista")}
                {validarCampo("NombreCompaniaTransportista")}
                {validarCampo("NumeroViaje")}
                {validarCampo("Conductor")}
                {validarCampo("DocumentoTransporte")}
                {validarCampo("Ficha")}
                {validarCampo("Placa")}
                {validarCampo("RutaTransporte")}
                {validarCampo("ZonaTransporte")}
                {validarCampo("NumeroAlbaran")}
            }},";

            return Transporte;
        }
        private static string InformacionesAdicionales()
        {
            for (int i = 70; i <= 91; i++)
            {
                if (dr[i].ToString() == "#e")
                {
                    if (i == 91) return "";
                }
                else
                {
                    break;
                }
            }
            string Comprador = $@"
            ""InformacionesAdicionales"":{{
                {validarCampo("FechaEmbarque")}
                {validarCampo("NumeroEmbarque")}
                {validarCampo("NumeroContenedor")}
                {validarCampo("NumeroReferencia")}
                {validarCampo("NombrePuertoEmbarque")}
                {validarCampo("CondicionesEntrega")}
                {validarCampo("TotalFob")}
                {validarCampo("Seguro")}
                {validarCampo("Flete")}
                {validarCampo("OtrosGastos")}
                {validarCampo("TotalCif")}
                {validarCampo("RegimenAduanero")}
                {validarCampo("NombrePuertoSalida")}
                {validarCampo("NombrePuertoDesembarque")}
                {validarCampo("PesoBruto")}
                {validarCampo("PesoNeto")}
                {validarCampo("UnidadPesoBruto")}
                {validarCampo("UnidadPesoNeto")}
                {validarCampo("CantidadBulto")}
                {validarCampo("UnidadBulto")}
                {validarCampo("VolumenBulto")}
                {validarCampo("UnidadVolumen")}
            }},";

            return Comprador;
        }
        private static string Comprador()
        {
            if (dr["RNCComprador"].ToString() == "#e" && dr["IdentificadorExtranjero"].ToString() == "#e") return "";
            string Comprador = $@"
            ""Comprador"":{{
                {validarCampo("RNCComprador")}
                {validarCampo("IdentificadorExtranjero")}
                {validarCampo("RazonSocialComprador")}
                {validarCampo("ContactoComprador")}
                {validarCampo("CorreoComprador")}
                {validarCampo("DireccionComprador")}
                {validarCampo("MunicipioComprador")}
                {validarCampo("ProvinciaComprador")}
                {validarCampo("PaisComprador")}
                {validarCampo("FechaEntrega")}
                {validarCampo("ContactoEntrega")}
                {validarCampo("DireccionEntrega")}
                {validarCampo("TelefonoAdicional")}
                {validarCampo("FechaOrdenCompra")}
                {validarCampo("NumeroOrdenCompra")}
                {validarCampo("CodigoInternoComprador")}
                {validarCampo("ResponsablePago")}
                {validarCampo("InformacionAdicionalComprador")}
            }},";

            return Comprador;
        }
        private static string TablaTelefonoEmisor()
        {

            if (dr["TelefonoEmisor[1]"].ToString() == "#e") return "";
            string TablaFormasPago = $@"
                ""TablaTelefonoEmisor"": {{
                    ""TelefonoEmisor"" : [";

            for (int i = 1; i <= 3; i++)
            {
                if (dr[$"TelefonoEmisor[{i}]"].ToString() != "#e")
                {
                    TablaFormasPago += $@"
                    ""{dr[$"TelefonoEmisor[{i}]"].ToString()}"",";
                }
            }

            TablaFormasPago += $@"]
                }},";
            return TablaFormasPago.Replace(@""",]", @"""]");
        }
        private static string TablaFormasPago()
        {
            if (dr["MontoPago[1]"].ToString() == "#e") return "";
            string TablaFormasPago = $@"
                ""TablaFormasPago"": {{
                    ""FormaDePago"" : [";

            for (int i = 1; i <= 7; i++)
            {
                if (dr[$"MontoPago[{i}]"].ToString() != "#e")
                {
                    TablaFormasPago += $@"
                    {{
                        ""FormaPago"": {dr[$"FormaPago[{i}]"].ToString()},	
                        ""MontoPago"": ""{dr[$"MontoPago[{i}]"].ToString()}""
                    }},";
                }
            }

            TablaFormasPago += $@"]
                }},";
            return TablaFormasPago.Replace("},]", "}]");
        }
        private static string validarCampo(string Campo, string caption = null)
        {
            caption = caption ?? Campo;
            if (dr[Campo].ToString() != "#e") return $@"""{caption}"": ""{dr[Campo].ToString()}"",";
            else return "";
        }
        public static DataTable ConvertExcelToDataTable(string filePath)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            FileInfo fileInfo = new FileInfo(filePath);

            using (var package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet hoja = package.Workbook.Worksheets[0];

                DataTable dataTable = new DataTable();

                // Encabezados
                for (int col = 1; col <= hoja.Dimension.End.Column; col++)
                {
                    dataTable.Columns.Add(hoja.Cells[1, col].Text.TrimEnd());
                }
                //FIlas
                int index = 0;
                for (int fila = 2; fila <= hoja.Dimension.End.Row; fila++, index++)
                {
                    DataRow row = dataTable.NewRow();
                    for (int col = 1; col <= hoja.Dimension.End.Column; col++)
                    {
                        row[col - 1] = hoja.Cells[fila, col].Text;
                    }
                    CasoDePrueba.Add(index, hoja.Cells[fila, 1].Text);
                    dataTable.Rows.Add(row);
                }



                return dataTable;
            }
        }
    }
}
