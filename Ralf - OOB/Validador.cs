using GestaoOperacoes.FidcRevenda.SFTP;
using GestaoOperacoes.FidcRevenda.SFTPBot;
using GestaoOperacoes.Models;
using GestaoOperacoes.Models.FidcRevenda;

//using Ionic.Zip;
using GestaoOperacoes.Sefaz;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Hosting;
using System.Xml;

namespace GestaoOperacoes.FidcRevenda.SFTPVerificador
{
    public class Validador
    {
        private Thread Processa = null;
        private string Estado = "Aguardando";
        private string Trabalho = "";
        private HttpContext Context = null;
        private string Errors = "";
        private string StackErrors = "";
        private IDictionary<string, string[]> dicRevendas = new Dictionary<string, string[]>();
        public static Validador Current { get; private set; } = null;
        private string localtempPath = $"{HostingEnvironment.ApplicationPhysicalPath}/Upload/testeFidcRevenda/Xml/";
        private string localzipPath = $"{HostingEnvironment.ApplicationPhysicalPath}/Upload/testeFidcRevenda/Zip/";
        private string localexcelPath = $"{HostingEnvironment.ApplicationPhysicalPath}/Upload/testeFidcRevenda/Excel/";
        private string nomearquivo = "";
        private Validador()
        {
            Current = this;
            IniciarThread();
            this.Context = HttpContext.Current;
        }

        public static void Iniciar()
        {
            if (Current == null)
            {
                Current = new Validador();
            }
            else if (Current.Processa == null || !Current.Processa.IsAlive)
            {
                Current.IniciarThread();
            }
        }

        private void IniciarThread()
        {
            Thread Zipar = new Thread(Current.Processar);
            Current.Processa = Zipar;
            Zipar.Start();
        }

        private void Processar()
        {
            HttpContext.Current = this.Context;
            //int ContadorExecs = 0;
            Errors = "";
            while (true)
            {
                try
                {
                    //Thread.Sleep(9000);
                    Trabalho = "";
                    Errors = "";
                    int processo = 001; //fazer função pegar processo
                    nomearquivo = processo.ToString() + "-" + Utils.Utils.GerarToken();
                    string zipPath = localzipPath + nomearquivo + ".zip";

                    var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create);
                    /*
                    using (FidcSFTP sftp = new FidcSFTP())
                    {
                        var filesFTP = sftp.GetFolderFiles();
                        if (filesFTP != null)
                        {
                            foreach (var file in filesFTP)
                            {
                                //FarmGerenciadorDeProcessoDeArquivos.AdicionarArquivo(file);
                                FidcRevendaGerenciadorDeProcessoDeArquivos.AdicionarArquivo(file, processo);
                            }
                        }
                    }*/
                    int quantidade = 0;
                    string[] files = Directory.GetFiles(localtempPath, "*.xml");
                    foreach (var file in files)
                    {
                        // Add the entry for each file
                        zip.CreateEntryFromFile(file, Path.GetFileName(file), CompressionLevel.Optimal);
                        quantidade = +1;
                    }
                    zip.Dispose();
                    Validar( quantidade);
                    /*
                        if (Validar(localtempPath, quantidade))
                    {
                        // salvar o excel no ft
                        string excelfile = localexcelPath + codigo + ".xls";
                        using (FidcSFTP sftp = new FidcSFTP())
                        {
                            FileStream Excel = File.OpenRead(excelfile);
                            DadosExcel excel = new DadosExcel()
                            {
                                Data = DateTime.Now,
                                Caminho = excelfile,
                                Processo = processo,
                            };

                            sftp.UploadProcessed(Excel, codigo + ".xls");
                            //upa no banco
                            using (AfortContext db = new AfortContext())
                            {   
                                db.FidcRevenda_DadosExcel.Add(excel);
                                db.SaveChanges();
                            }
                        }
                    }
                    */
                    Console.WriteLine("Done!");
                }
                catch (Exception e)
                {
                    Errors = e.Message + "\n";
                    StackErrors = e.StackTrace;
                }
                Estado = "Aguardando";

                Thread.Sleep(1800000);
            }
        }

        public string EstadoAtual()
        {
            return Estado;
        }

        public string TrabalhoAtual()
        {
            return Trabalho;
        }

        public string GerErrors()
        {
            return Errors.Replace("\n", "<br />");
        }

        public Boolean Validar(int quantrows)
        {
            if (localtempPath != null)
            {
                // string UrlAPI = "https://nfe.fazenda.mg.gov.br/nfe2/services/NFeConsultaProtocolo4";

                // string action = "nfeConsultaNF";
                Dictionary<string, string> xmlRequest = new Dictionary<string, string>();

                //string xmlconsulta = Utils.SOAP_Request.SendSOAPRequest(UrlAPI, action, xmlRequest, "nfeConsultaNF", true);
                ServicosSefaz servicosSefaz = new ServicosSefaz();

                string documento = "";

                //using (FileStream criarArquivo = new FileStream(zpath, FileMode.Create))
                //{
                //    zpath.InputStream.CopyTo(criarArquivo);
                //}

                //string nomezip = "";

                //foreach (FileInfo file in dirInfo.GetFiles())
                //{
                //    nomezip += file.FullName;
                //}
                DirectoryInfo dirInfo = new DirectoryInfo(localtempPath);
                //System.IO.File.Delete(nomezip);

                bool sucesso = false;
                string marcador = "";

                // int quantrows = dirInfo.GetFiles().Length;

                string header = "<?xml version=\"1.0\"?>" +
                                   "<?mso-application progid=\"Excel.Sheet\"?>" +
                                   "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"" +
                                   " xmlns:o=\"urn:schemas-microsoft-com:office:office\"" +
                                   " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"" +
                                   " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"" +
                                   " xmlns:html=\"http://www.w3.org/TR/REC-html40\">" +
                                   " <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">" +
                                   "  <Author>Desenvolvimento</Author>" +
                                   "  <LastAuthor>Desenvolvimento</LastAuthor>" +
                                   "  <Created>2015-06-05T18:19:34Z</Created>" +
                                   "  <LastSaved>2019-10-14T14:27:52Z</LastSaved>" +
                                   "  <Version>16.00</Version>" +
                                   " </DocumentProperties>" +
                                   " <OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\">" +
                                   "  <AllowPNG/>" +
                                   " </OfficeDocumentSettings>" +
                                   " <ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">" +
                                   "  <WindowHeight>7485</WindowHeight>" +
                                   "  <WindowWidth>20490</WindowWidth>" +
                                   "  <WindowTopX>32767</WindowTopX>" +
                                   "  <WindowTopY>32767</WindowTopY>" +
                                   "  <ProtectStructure>False</ProtectStructure>" +
                                   "  <ProtectWindows>False</ProtectWindows>" +
                                   " </ExcelWorkbook>" +
                               " <Styles>" +
                               "  <Style ss:ID=\"Default\" ss:Name=\"Normal\">" +
                               "   <Alignment ss:Vertical=\"Bottom\"/>" +
                               "   <Borders/>" +
                               "   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/>" +
                               "   <Interior/>" +
                               "   <NumberFormat/>" +
                               "   <Protection/>" +
                               "  </Style>" +
                               "  <Style ss:ID=\"s62\">" +
                               "   <Interior ss:Color=\"#FFFF00\" ss:Pattern=\"Solid\"/>" +
                               "  </Style>" +
                               "  <Style ss:ID=\"s63\">" +
                               "   <Borders>" +
                               "    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" +
                               "   </Borders>" +
                               "   <Interior ss:Color=\"#FFFF00\" ss:Pattern=\"Solid\"/>" +
                               "  </Style>" +
                                "  <Style ss:ID=\"s64\">" +
                               "   <NumberFormat ss:Format=\"Short Date\"/>" +
                               "  </Style>" +
                               "  <Style ss:ID=\"s65\">" +
                               "   <NumberFormat ss:Format=\"&quot;R$&quot;\\ #,##0.00\"/>" +
                               "  </Style>" +
                               "  <Style ss:ID=\"s66\">" +
                               "   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/>" +
                               "   <Interior ss:Color=\"#E2EFDA\" ss:Pattern=\"Solid\"/>" +
                               "  </Style>" +
                               "<Style ss:ID=\"s68\">" +
                               "   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#FF0000\"/>" +
                               "   <Interior ss:Color=\"#F4B084\" ss:Pattern=\"Solid\"/>" +
                               "  </Style>" +
                               "<Style ss:ID=\"s70\">" +
                               "   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#375623\"/>" +
                               "   <Interior ss:Color=\"#E2EFDA\" ss:Pattern=\"Solid\"/>" +
                               "  </Style>" +
                               " </Styles>" +
                               " <Worksheet ss:Name=\"Plan1\">" +
                               "  <Names>" +
                               "   <NamedRange ss:Name=\"_FilterDatabase\" ss:RefersTo=\"=Plan1!R1C1:R1C50\"" +
                               "    ss:Hidden=\"1\"/>" +
                               "  </Names>" +
                               "  <Table ss:ExpandedColumnCount=\"57\" ss:ExpandedRowCount=\"" + quantrows + 1 + "\" x:FullColumns=\"1\"" +
                               "   x:FullRows=\"1\" ss:DefaultRowHeight=\"15\">" +
                               "<Column ss:AutoFitWidth=\"0\" ss:Width=\"288.75\"/>" +
                               "<Column ss:AutoFitWidth=\"0\" ss:Width=\"120.75\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"68.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"120.75\"/>" +
                                "   <Column ss:Index=\"5\" ss:AutoFitWidth=\"0\" ss:Width=\"96.75\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"96.75\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"108\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"44.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"93.75\"/>" +
                                "   <Column ss:Index=\"10\" ss:AutoFitWidth=\"0\" ss:Width=\"63\"/>" +
                                "   <Column ss:Index=\"13\" ss:AutoFitWidth=\"0\" ss:Width=\"106.5\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"83.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"122.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"117\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"87\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"95.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"83.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"150\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"60.75\"/>" +
                                "   <Column ss:Index=\"23\" ss:AutoFitWidth=\"0\" ss:Width=\"64.5\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"70.5\"/>" +
                                "   <Column ss:Index=\"26\" ss:AutoFitWidth=\"0\" ss:Width=\"60.75\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"72\"/>" +
                                "   <Column ss:Index=\"29\" ss:AutoFitWidth=\"0\" ss:Width=\"77.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"69\"/>" +
                                "   <Column ss:Index=\"32\" ss:AutoFitWidth=\"0\" ss:Width=\"79.5\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"64.5\"/>" +
                                "   <Column ss:Index=\"35\" ss:AutoFitWidth=\"0\" ss:Width=\"62.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"71.25\"/>" +
                                "   <Column ss:Index=\"38\" ss:AutoFitWidth=\"0\" ss:Width=\"58.5\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"61.5\"/>" +
                                "   <Column ss:Index=\"41\" ss:AutoFitWidth=\"0\" ss:Width=\"66.75\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"67.5\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"60.75\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"58.5\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"59.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"57.75\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"65.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"62.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"60\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"66.75\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"68.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"150\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"68.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"68.25\"/>" +
                                "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"68.25\"/>" +
                                "<Row ss:AutoFitHeight=\"0\">" +
                                "    <Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">Nome Arquivo</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                  "    <Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">Revenda</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">Tipo Docto</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                 "    <Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">Nº Docto (CPF ou CNPJ)</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">Calculo</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">Nome Devedor</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Rua</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "<Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Nº</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Complemento</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Bairro</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                 "<Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Cidade</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">UF</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">CEP</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Tipo Recebivel</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Numero NF</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Natureza da Operação</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Forma de Pagamento</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Chave</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Data Emissao</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor NF</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Validado</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Status</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "<Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-1</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-1</Data><NamedCell" +
                                "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-1</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-2</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-2</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-2</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-3</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-3</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-3</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-4</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-4</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-4</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-5</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-5</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-5</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-6</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-6</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-6</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-7</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-7</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-7</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-8</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-8</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-8</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-9</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-9</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-9</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Duplicata-10</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Vencto-10</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Valor-10</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Informação Complementar</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                         "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">UF Emitente</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                         "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Inscrição Estadual Emitente</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                         "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Inscrição Estadual Destinatário</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                        "    <Cell ss:StyleID=\"s63\"><Data ss:Type=\"String\">Codigo</Data><NamedCell" +
                                        "      ss:Name=\"_FilterDatabase\"/></Cell>" +
                                "   </Row>";

                string footer = "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">" +
                                                "   <PageSetup>" +
                                                "    <Header x:Margin=\"0.3\"/>" +
                                                "    <Footer x:Margin=\"0.3\"/>" +
                                                "    <PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/>" +
                                                "   </PageSetup>" +
                                                "   <Unsynced/>" +
                                                "   <Print>" +
                                                "    <ValidPrinterInfo/>" +
                                                "    <PaperSizeIndex>9</PaperSizeIndex>" +
                                                "    <HorizontalResolution>300</HorizontalResolution>" +
                                                "    <VerticalResolution>300</VerticalResolution>" +
                                                "   </Print>" +
                                                "   <Selected/>" +
                                                "   <FreezePanes/>" +
                                                "   <FrozenNoSplit/>" +
                                                "   <SplitHorizontal>1</SplitHorizontal>" +
                                                "   <TopRowBottomPane>1</TopRowBottomPane>" +
                                                "   <ActivePane>2</ActivePane>" +
                                                "   <Panes>" +
                                                "    <Pane>" +
                                                "     <Number>3</Number>" +
                                                "     <ActiveCol>1</ActiveCol>" +
                                                "    </Pane>" +
                                                "    <Pane>" +
                                                "     <Number>2</Number>" +
                                                "     <ActiveRow>14</ActiveRow>" +
                                                "     <ActiveCol>3</ActiveCol>" +
                                                "    </Pane>" +
                                                "   </Panes>" +
                                                "   <ProtectObjects>False</ProtectObjects>" +
                                                "   <ProtectScenarios>False</ProtectScenarios>" +
                                                "  </WorksheetOptions>" +
                                                "  <AutoFilter x:Range=\"R1C1:R1C54\"" +
                                                "   xmlns=\"urn:schemas-microsoft-com:office:excel\">" +
                                                "  </AutoFilter>" +
                                                " </Worksheet>" +
                                                "</Workbook>";

                documento += header;

                XmlTextReader reader = null;
                string revenda = "";
                string confConsultar = System.Configuration.ConfigurationManager.AppSettings["ConsultaSEFAZ"];
                bool consultar = true;
                List<NotaFiscal> notas = new List<NotaFiscal>();

                foreach (FileInfo file in dirInfo.GetFiles())
                {
                    NotaFiscal nota = new NotaFiscal();

                    XmlDocument xml = new XmlDocument();
                    xml.PreserveWhitespace = true;

                    nota.Nomearquivo = file.Name;

                    try
                    {
                        xml.Load(file.FullName);
                    }
                    catch (Exception e)
                    {
                        nota.Nomearquivo = file.Name + " - Arquivo inválido.";
                    }

                    if (xml.GetElementsByTagName("NFe").Count != 0)
                    {
                        sucesso = Utils.NotaFiscal.CheckSignatureNFe(xml.GetElementsByTagName("NFe"));
                    }

                    if (sucesso)
                    {
                        marcador = "Assinatura válida.";
                    }
                    else
                    {
                        marcador = "Assinatura inválida.";
                    }

                    XmlNamespaceManager xmlnsManager = new XmlNamespaceManager(xml.NameTable);
                    xmlnsManager.AddNamespace("nfe", "http://www.portalfiscal.inf.br/nfe");

                    XmlNodeList NomeDevedor = xml.SelectNodes("//nfe:dest/nfe:xNome", xmlnsManager);
                    if (NomeDevedor.Count > 0)
                    {
                        XmlNode Nome = NomeDevedor[0];
                        nota.Nome_Destinatario = Nome.InnerText;
                    }

                    XmlNodeList CpfDevedor = xml.SelectNodes("//nfe:dest/nfe:CPF", xmlnsManager);
                    if (CpfDevedor.Count > 0)
                    {
                        XmlNode Cpf = CpfDevedor[0];
                        nota.Documento_Destinatario = Utils.Utils.formatarpraCPF(Cpf.InnerText);
                    }

                    XmlNodeList CNPJDevedor = xml.SelectNodes("//nfe:dest/nfe:CNPJ", xmlnsManager);
                    if (CNPJDevedor.Count > 0)
                    {
                        XmlNode Cnpj = CNPJDevedor[0];
                        nota.Documento_Destinatario = Utils.Utils.formatarpraCNPJ(Cnpj.InnerText);
                    }

                    if (nota.Documento_Destinatario != null && nota.Documento_Destinatario.Length == 14)
                    {
                        nota.TipoDocto = "PF";
                    }
                    else if (nota.Documento_Destinatario != null && nota.Documento_Destinatario.Length == 18)
                    {
                        nota.TipoDocto = "PJ";
                    }

                    XmlNodeList LogradouroDevedor = xml.SelectNodes("//nfe:dest/nfe:enderDest/nfe:xLgr", xmlnsManager);
                    if (LogradouroDevedor.Count > 0)
                    {
                        XmlNode Logradouro = LogradouroDevedor[0];
                        nota.Logradouro_Destinatario = Logradouro.InnerText;
                    }

                    XmlNodeList NLogradouroDevedor = xml.SelectNodes("//nfe:dest/nfe:enderDest/nfe:nro", xmlnsManager);
                    if (NLogradouroDevedor.Count > 0)
                    {
                        XmlNode NLogradouro = NLogradouroDevedor[0];
                        nota.NLogradouro_Destinatario = NLogradouro.InnerText;
                    }

                    XmlNodeList ComplementoDevedor = xml.SelectNodes("//nfe:dest/nfe:enderDest/nfe:xCpl", xmlnsManager);
                    if (ComplementoDevedor.Count > 0)
                    {
                        XmlNode Complemento = ComplementoDevedor[0];
                        nota.Complemento_Destinatario = Complemento.InnerText;
                    }

                    XmlNodeList BairroDevedor = xml.SelectNodes("//nfe:dest/nfe:enderDest/nfe:xBairro", xmlnsManager);
                    if (BairroDevedor.Count > 0)
                    {
                        XmlNode Bairro = BairroDevedor[0];
                        nota.Bairro_Destinatario = Bairro.InnerText;
                    }

                    XmlNodeList CidadeDevedor = xml.SelectNodes("//nfe:dest/nfe:enderDest/nfe:xMun", xmlnsManager);
                    if (CidadeDevedor.Count > 0)
                    {
                        XmlNode Cidade = CidadeDevedor[0];
                        nota.Cidade_Destinatario = Cidade.InnerText;
                    }

                    XmlNodeList UFDevedor = xml.SelectNodes("//nfe:dest/nfe:enderDest/nfe:UF", xmlnsManager);
                    if (UFDevedor.Count > 0)
                    {
                        XmlNode UF = UFDevedor[0];
                        nota.UF_Destinatario = UF.InnerText;
                    }

                    XmlNodeList CEPDevedor = xml.SelectNodes("//nfe:dest/nfe:enderDest/nfe:CEP", xmlnsManager);
                    if (CEPDevedor.Count > 0)
                    {
                        XmlNode CEP = CEPDevedor[0];
                        nota.CEP_Destinatario = CEP.InnerText;
                    }

                    XmlNodeList NumeroNF = xml.SelectNodes("//nfe:ide/nfe:nNF", xmlnsManager);
                    if (NumeroNF.Count > 0)
                    {
                        XmlNode Numero = NumeroNF[0];
                        nota.NF = Numero.InnerText;
                    }

                    XmlNodeList NaturezaOperacao = xml.SelectNodes("//nfe:ide/nfe:natOp", xmlnsManager);
                    if (NaturezaOperacao.Count > 0)
                    {
                        XmlNode NaturezaOp = NaturezaOperacao[0];
                        nota.NaturezaOperacao = NaturezaOp.InnerText;
                    }

                    XmlNodeList FormaPagamento = xml.SelectNodes("//nfe:ide/nfe:indPag", xmlnsManager);
                    if (FormaPagamento.Count > 0)
                    {
                        XmlNode Pagamento = FormaPagamento[0];
                        nota.FormaPagamento = Pagamento.InnerText;
                        if (nota.FormaPagamento == "0")
                        {
                            nota.FormaPagamento = "À vista";
                        }
                        else if (nota.FormaPagamento == "1")
                        {
                            nota.FormaPagamento = "À prazo";
                        }
                        else
                        {
                            nota.FormaPagamento = "Inválido";
                        }
                    }

                    XmlNodeList ChaveNota = xml.SelectNodes("//nfe:infNFe", xmlnsManager);
                    if (ChaveNota.Count > 0)
                    {
                        XmlNode Chave = ChaveNota[0];
                        nota.ChavedeAcesso = Chave.Attributes["Id"].Value;
                    }

                    //informação adicional
                    XmlNodeList infCpls = xml.SelectNodes("//nfe:infCpl", xmlnsManager);
                    if (infCpls.Count > 0)
                    {
                        XmlNode infCpl = infCpls[0];
                        nota.InfCpl = infCpl.InnerText;
                    }

                    XmlNodeList UFEmitente = xml.SelectNodes("//nfe:emit/nfe:enderEmit/nfe:UF", xmlnsManager);
                    if (UFEmitente.Count > 0)
                    {
                        XmlNode UF = UFEmitente[0];
                        nota.UF_Emitente = UF.InnerText;
                    }

                    XmlNodeList IEEmitente = xml.SelectNodes("//nfe:emit/nfe:IE", xmlnsManager);
                    if (IEEmitente.Count > 0)
                    {
                        XmlNode IEEmit = IEEmitente[0];
                        nota.IEEmitente = IEEmit.InnerText;
                    }

                    XmlNodeList IEDestinatario = xml.SelectNodes("//nfe:dest/nfe:IE", xmlnsManager);
                    if (IEDestinatario.Count > 0)
                    {
                        XmlNode IEDest = IEDestinatario[0];
                        nota.IEDestinatario = IEDest.InnerText;
                    }

                    ///////SEFAZ

                    string statusCode = "0";
                    string message = "Falha de comunicação com o servidor";
                    string resultadoxml;
                    if (consultar)
                    {
                        try
                        {
                            resultadoxml = servicosSefaz.NfeConsultaProtocolo(nota.ChavedeAcesso.ToLower().Replace("nfe", ""));
                            XmlDocument r = new XmlDocument();
                            r.LoadXml(resultadoxml);
                            XmlNodeList cStat = r.GetElementsByTagName("cStat");
                            XmlNodeList xMotivo = r.GetElementsByTagName("xMotivo");

                            if (cStat.Count > 0 && xMotivo.Count > 0)
                            {
                                statusCode = cStat[0].InnerText;
                                message = xMotivo[0].InnerText;
                            }
                        }
                        catch (Exception e)
                        {
                            message += " - " + e.Message;
                        }
                    }
                    else
                    {
                        message = "Consulta Desativada";
                    }
                    nota.Status = message;
                    nota.StatusCode = statusCode;
                    ///////
                    //////

                    XmlNodeList DataEmissao = xml.SelectNodes("//nfe:ide/nfe:dhEmi", xmlnsManager);
                    if (DataEmissao.Count > 0)
                    {
                        XmlNode Data = DataEmissao[0];
                        nota.Data_Emissao = Data.InnerText;
                    }

                    XmlNodeList ValorDevedor = xml.SelectNodes("//nfe:total/nfe:ICMSTot/nfe:vNF", xmlnsManager);
                    if (ValorDevedor.Count > 0)
                    {
                        XmlNode Valor = ValorDevedor[0];
                        nota.ValorNF = Valor.InnerText;
                    }

                    XmlNodeList numerodup = xml.SelectNodes("//nfe:cobr/nfe:dup/nfe:nDup", xmlnsManager);

                    List<string> numeroduplicatas = new List<string>();

                    if (numerodup.Count > 0)
                    {
                        for (int i = 0; i < numerodup.Count; i++)
                        {
                            if (numeroduplicatas.Count < 10)
                            {
                                XmlNode numeroduplicata = numerodup[i];
                                numeroduplicatas.Add(numeroduplicata.InnerText);
                            }
                        }
                    }

                    XmlNodeList dVenc = xml.SelectNodes("//nfe:cobr/nfe:dup/nfe:dVenc", xmlnsManager);

                    List<string> dVencDuplicatas = new List<string>();

                    if (dVenc.Count > 0)
                    {
                        for (int i = 0; i < dVenc.Count; i++)
                        {
                            if (dVencDuplicatas.Count < 10)
                            {
                                XmlNode dVencduplicata = dVenc[i];
                                dVencDuplicatas.Add(dVencduplicata.InnerText);
                            }
                        }
                    }

                    XmlNodeList vDup = xml.SelectNodes("//nfe:cobr/nfe:dup/nfe:vDup", xmlnsManager);

                    List<string> vDuplicatas = new List<string>();

                    if (vDup.Count > 0)
                    {
                        for (int i = 0; i < vDup.Count; i++)
                        {
                            if (vDuplicatas.Count < 10)
                            {
                                XmlNode vduplicata = vDup[i];
                                vDuplicatas.Add(vduplicata.InnerText);
                            }
                        }
                    }

                    List<DuplicatasXml> duplicatas = new List<DuplicatasXml>();

                    for (int i = 0; i < numeroduplicatas.Count; i++)
                    {
                        DuplicatasXml novaduplicata = new DuplicatasXml();
                        novaduplicata.nDup = numeroduplicatas[i];
                        novaduplicata.vDup = vDuplicatas[i];
                        novaduplicata.dVenc = dVencDuplicatas[i];

                        duplicatas.Add(novaduplicata);
                    }

                    nota.Validado = marcador;

                    nota.TipoRecebivel = "NF";

                    XmlNodeList emitente = xml.SelectNodes("//nfe:emit/nfe:xNome", xmlnsManager);
                    if (emitente.Count > 0)
                    {
                        XmlNode Valor = emitente[0];
                        nota.Revenda = Valor.InnerText;
                    }

                    //revenda = file.FullName;
                    //string nomearq = file.Name;

                    //if (revenda.Contains("Empresa a"))
                    //{
                    //    nota.Revenda = "Revenda A";
                    //} else if (revenda.Contains("Empresa B"))
                    //{
                    //    nota.Revenda = "Revenda B";
                    //}
                    //switch (expression)
                    //{
                    //    case x:
                    //        // code block
                    //        break;
                    //    case y:
                    //        // code block
                    //        break;
                    //    default:
                    //        // code block
                    //        break;
                    //}

                    nota.duplicatas = duplicatas;
                    notas.Add(nota);

                    System.IO.File.Delete(file.FullName);
                    sucesso = false;
                }

                documento += Utils.NotaFiscal.gerarCorpo(notas);
                documento += "</Table>";
                documento += footer;

                byte[] excel = Encoding.UTF8.GetBytes(documento);

                string nomeexcel = nomearquivo + ".xls";
                string salvarExcel = localexcelPath + nomeexcel;

                using (FileStream criarArquivoExcel = new FileStream(salvarExcel, FileMode.Create))
                {
                    criarArquivoExcel.Write(excel, 0, excel.Length);
                }

                //System.IO.Directory.Delete(temppath);

                // System.IO.File.Delete(Nome);

                return true;

                //if (System.IO.File.Exists(zpath))
                //{
                //    System.IO.Directory.Delete(zpath);
                //}
            }
            else
            {
                return false;
            }
        }
    }
}