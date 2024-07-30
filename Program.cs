using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace GeradorRelatorioPDF
{
    class Program
    {
        static List<Pessoa> pessoas = new List<Pessoa>();
        static BaseFont fonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);

        static void Main(string[] args)
        {

            DesserializarPessoas();
            //foreach (var p in pessoas)
            //{
            //    Console.WriteLine($"{p.IdPessoa} - {p.Nome} {p.Sobrenome}");
            //}
            GerarRelatorioPDF(186);
        }

        static void DesserializarPessoas()
        {
            if (File.Exists("pessoas.json"))
            {
                using (var sr = new StreamReader("pessoas.json"))
                {
                    var dados = sr.ReadToEnd();
                    pessoas = JsonSerializer.Deserialize(dados, typeof(List<Pessoa>)) as List<Pessoa>;
                }
            }
        }

        static void GerarRelatorioPDF(int qtdePessoas)
        {
            var pessoasSelecionadas = pessoas.Take(qtdePessoas).ToList();
            if(pessoasSelecionadas.Count > 0)
            {

                // Calculo da quantidade total de paginas (consigerando valores fixos de row)
                int totalPaginas = 1;
                int totalLinhas = pessoasSelecionadas.Count;
                if (totalLinhas > 29)
                {
                    totalPaginas += (int)Math.Ceiling((totalLinhas - 29) / 30F);
                }

                // Configuração do documento PDF
                var pxPorMm = 72 / 32.2F;
                var pdf = new Document(PageSize.A4, 15 * pxPorMm, 15 * pxPorMm, 15 * pxPorMm, 20 * pxPorMm);
                var nomeArquivo = $"pessoas.{DateTime.Now.ToString("yyyy.MM.dd.hh.mm.ss")}.pdf";
                var arquivo = new FileStream(nomeArquivo, FileMode.Create);
                var writer = PdfWriter.GetInstance(pdf, arquivo);
                writer.PageEvent = new EventosDePagina(totalPaginas);

                pdf.Open();

                // Adição do titulo
                var fontParagrafo = new iTextSharp.text.Font(fonteBase, 22, iTextSharp.text.Font.NORMAL, BaseColor.Black);
                var fontParagrafo2 = new iTextSharp.text.Font(fonteBase, 10, iTextSharp.text.Font.ITALIC, BaseColor.Black);
                var titulo = new Paragraph("Relatório de Pessoas\n", fontParagrafo);
                var titulo2 = new Paragraph("As empresas costumam fazer consultas a um CPF quando estão prestes a fechar\n um negócio ou quando o consumidor solicita crédito. \n\n", fontParagrafo2);
                titulo.Alignment = Element.ALIGN_LEFT;
                titulo.SpacingAfter = 4;
                pdf.Add(titulo);
                pdf.Add(titulo2);

                // Adião da imagem
                var caminhoImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img\\youtube.png");
                if (File.Exists(caminhoImage))
                {
                    iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImage);
                    float razaoAlturaLargura = logo.Width / logo.Height;
                    float alturaLogo = 75;
                    float larguraLogo = alturaLogo * razaoAlturaLargura;
                    logo.ScaleToFit(larguraLogo, alturaLogo);
                    var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                    var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 55;
                    logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                    writer.DirectContent.AddImage(logo, false);
                }

                // Adição da tabela de dados
                var tabela = new PdfPTable(5);
                float[] largurasColunas = { 0.6f, 2f, 1.5f, 1f, 1f };
                tabela.SetWidths(largurasColunas);
                tabela.DefaultCell.BorderWidth = 0;
                tabela.WidthPercentage = 100;

                // Adição da celulas de titulos das colunas
                CriarCelulaTexto(tabela, "Código", PdfCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Nome", PdfCell.ALIGN_LEFT, true);
                CriarCelulaTexto(tabela, "Profissão", PdfCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Salário", PdfCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Empregada", PdfCell.ALIGN_CENTER, true);

                // Adição dos valores das celulas
                foreach(var p in pessoasSelecionadas)
                {
                    CriarCelulaTexto(tabela, p.IdPessoa.ToString("D6"), PdfPCell.ALIGN_CENTER);
                    CriarCelulaTexto(tabela, p.Nome + " " + p.Sobrenome);
                    CriarCelulaTexto(tabela, p.Profissao.Nome, PdfPCell.ALIGN_CENTER);
                    CriarCelulaTexto(tabela, p.Salario.ToString("C2"), PdfPCell.ALIGN_RIGHT);
                    //CriarCelulaTexto(tabela, p.Empregado ? "Sim" : "Não", PdfPCell.ALIGN_CENTER);
                    var caminhoImagemCelula = p.Empregado ? "img\\emoji_feliz.png" : "img\\emoji_triste.png";
                    caminhoImagemCelula = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, caminhoImagemCelula);
                    CriarCelulaImagem(tabela, caminhoImagemCelula, 20, 20);
                }


                pdf.Add(tabela);

                pdf.Close();
                arquivo.Close();

                // Abre o PDF no visualizador padrão
                //var caminhoPDF = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo);
                //if (File.Exists(caminhoPDF))
                //{
                //    Process.Start(new ProcessStartInfo()
                //    {
                //        Arguments = $"/c start {caminhoPDF}",
                //        FileName = "cmd.exe",
                //        CreateNoWindow = true
                //    }
                //    );
                //}
            }

        }

        static void CriarCelulaTexto(PdfPTable tabela, string texto, int alinhamentoHorizontal = PdfPCell.ALIGN_LEFT, bool negrito = false, bool italico = false, int tamanhoFonte = 12, int alturaCelula = 25 )
        {
            int estilo = iTextSharp.text.Font.NORMAL;
            if (negrito && italico)
            {
                estilo = iTextSharp.text.Font.BOLDITALIC;

            }else if (negrito)
            {
                estilo = iTextSharp.text.Font.BOLD;

            }
            else if (italico)
            {
                estilo = iTextSharp.text.Font.ITALIC;
            }

            var fonteCelula = new iTextSharp.text.Font(fonteBase, tamanhoFonte, estilo, BaseColor.Black);
            var bgColor = iTextSharp.text.BaseColor.White;

            if (tabela.Rows.Count % 2 == 1)
            {
                bgColor = new BaseColor(0.95F, 0.95F, 0.95F);
            }

            var celula = new PdfPCell(new Phrase(texto, fonteCelula));
            celula.HorizontalAlignment = alinhamentoHorizontal;
            celula.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            celula.Border = 0;
            celula.BorderWidthBottom = 1;
            celula.FixedHeight = alturaCelula;
            celula.PaddingBottom = 5;
            celula.BackgroundColor = bgColor;
            tabela.AddCell(celula);
        }

        static void CriarCelulaImagem(PdfPTable tabela, string caminhoImagem, int larguraImagem, int alturaImagem, int alturaCelula = 25)
        {
            var bgColor = iTextSharp.text.BaseColor.White;

            if (tabela.Rows.Count % 2 == 1)
            {
                bgColor = new BaseColor(0.95F, 0.95F, 0.95F);
            }

            if (File.Exists(caminhoImagem))
            {
                iTextSharp.text.Image imagem = iTextSharp.text.Image.GetInstance(caminhoImagem);
                imagem.ScaleToFit(larguraImagem, alturaImagem);

                var celula = new PdfPCell(imagem);
                celula.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                celula.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                celula.Border = 0;
                celula.BorderWidthBottom = 1;
                celula.FixedHeight = alturaCelula;
                celula.BackgroundColor = bgColor;
                tabela.AddCell(celula);
            }
        }
    }
}
