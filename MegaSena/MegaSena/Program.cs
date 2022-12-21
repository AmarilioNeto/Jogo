using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;


namespace MegaSena
{
    class Program
    {
        static int maior = 0;
        static void Main(string[] args)
        {
            Console.WriteLine("Seja bem vindo cole aqui o caminho em excel de todos os jogos da mega sena");
           var caminho = Console.ReadLine();         
            Workbook wb = new Workbook(caminho);
            Console.WriteLine("Escolha que tipo de Jogo irá fazer");
            Console.WriteLine("1) Com os números que mais sorteados");
            Console.WriteLine("2) Com os números que menos menos sorteados");
            Console.WriteLine("3) Com os números de jogos mais sorteados");
            Console.WriteLine("4) Com os números de jogos mais sorteados");
            var opcao = Console.ReadLine();
            List<int> listGeral = new List<int>();
           
            List<int> num1 = new List<int>();
            List<int> num2 = new List<int>();
            List<int> num3 = new List<int>();
            List<int> num4 = new List<int>();
            List<int> num5 = new List<int>();
            List<int> num6 = new List<int>();
            List<int> num7 = new List<int>();
            List<int> num8 = new List<int>();
            List<int> num9 = new List<int>();
            List<int> num10 = new List<int>();
            List<int> num11 = new List<int>();
            List<int> num12 = new List<int>();
            List<int> num13 = new List<int>();
            List<int> num14 = new List<int>();
            List<int> num15 = new List<int>();
            List<int> num16 = new List<int>();
            List<int> num17 = new List<int>();
            List<int> num18 = new List<int>();
            List<int> num19 = new List<int>();
            List<int> num20 = new List<int>();
            List<int> num21 = new List<int>();
            List<int> num22 = new List<int>();
            List<int> num23 = new List<int>();
            List<int> num24 = new List<int>();
            List<int> num25 = new List<int>();
            List<int> num26 = new List<int>();
            List<int> num27 = new List<int>();
            List<int> num28 = new List<int>();
            List<int> num29 = new List<int>();
            List<int> num30 = new List<int>();
            List<int> num31 = new List<int>();
            List<int> num32 = new List<int>();
            List<int> num33 = new List<int>();
            List<int> num34 = new List<int>();
            List<int> num35 = new List<int>();
            List<int> num36 = new List<int>();
            List<int> num37 = new List<int>();
            List<int> num38 = new List<int>();
            List<int> num39 = new List<int>();
            List<int> num40 = new List<int>();
            List<int> num41 = new List<int>();
            List<int> num42 = new List<int>();
            List<int> num43 = new List<int>();
            List<int> num44 = new List<int>();
            List<int> num45 = new List<int>();
            List<int> num46 = new List<int>();
            List<int> num47 = new List<int>();
            List<int> num48 = new List<int>();
            List<int> num49 = new List<int>();
            List<int> num50 = new List<int>();
            List<int> num51 = new List<int>();
            List<int> num52 = new List<int>();
            List<int> num53 = new List<int>();
            List<int> num54 = new List<int>();
            List<int> num55 = new List<int>();
            List<int> num56 = new List<int>();
            List<int> num57 = new List<int>();
            List<int> num58 = new List<int>();
            List<int> num59 = new List<int>();
            List<int> num60 = new List<int>();

            if (opcao == "1")
            {
                WorksheetCollection collection = wb.Worksheets;

                for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
                {

                    // Obter planilha usando seu índice
                    Worksheet worksheet = collection[worksheetIndex];                   
                    // Obter número de linhas e colunas
                    int rows = worksheet.Cells.MaxDataRow;
                    int cols = worksheet.Cells.MaxDataColumn;

                    // Percorrer as linhas
                    for (int i = 1; i < rows; i++)
                    {

                        // Percorrer cada coluna na linha selecionada
                        for (int j = 2; j > 1 && j < 8  ; j++)
                        {
                            
                            // Valor da célula de Pring
                            Console.Write(worksheet.Cells[i, j].Value + " | ");
                            for (int l = 1; l <= 60; l++)
                            {
                                if (Convert.ToInt32(worksheet.Cells[i, j].Value) == l)
                                {
                                    if (l == 1)
                                    {
                                        num1.Add(l);
                                        break;
                                    }
                                    if (l == 2)
                                    {
                                        num2.Add(l);
                                        break;
                                    }
                                    if (l == 3)
                                    {
                                        num3.Add(l);
                                        break;
                                    }
                                    if (l == 4)
                                    {
                                        num4.Add(l);
                                        break;
                                    }
                                    if (l == 5)
                                    {
                                        num5.Add(l);
                                        break;
                                    }
                                    if (l == 6)
                                    {
                                        num6.Add(l);
                                        break;
                                    }
                                    if (l == 7)
                                    {
                                        num7.Add(l);
                                        break;
                                    }
                                    if (l == 8)
                                    {
                                        num8.Add(l);
                                        break;
                                    }
                                    if (l == 9)
                                    {
                                        num9.Add(l);
                                        break;
                                    }
                                    if (l == 10)
                                    {
                                        num10.Add(l);
                                        break;
                                    }
                                    
                                    if (l == 11)
                                    {
                                        num11.Add(l);
                                        break;
                                    }
                                    if (l == 12)
                                    {
                                        num12.Add(l);
                                        break;
                                    }
                                    if (l == 13)
                                    {
                                        num13.Add(l);
                                        break;
                                    }
                                    if (l == 14)
                                    {
                                        num14.Add(l);
                                        break;
                                    }
                                    if (l == 15)
                                    {
                                        num15.Add(l);
                                        break;
                                    }
                                    if (l == 16)
                                    {
                                        num16.Add(l);
                                        break;
                                    }
                                    if (l == 17)
                                    {
                                        num17.Add(l);
                                        break;
                                    }
                                    if (l == 18)
                                    {
                                        num18.Add(l);
                                        break;
                                    }
                                    if (l == 19)
                                    {
                                        num19.Add(l);
                                        break;
                                    }
                                    if (l == 20)
                                    {
                                        num20.Add(l);
                                        break;
                                    }
                                    if (l == 21)
                                    {
                                        num21.Add(l);
                                        break;
                                    }
                                    if (l == 22)
                                    {
                                        num22.Add(l);
                                        break;
                                    }
                                    if (l == 23)
                                    {
                                        num23.Add(l);
                                        break;
                                    }
                                    if (l == 24)
                                    {
                                        num24.Add(l);
                                        break;
                                    }
                                    if (l == 25)
                                    {
                                        num25.Add(l);
                                        break;
                                    }
                                    if (l == 26)
                                    {
                                        num26.Add(l);
                                        break;
                                    }
                                    if (l == 27)
                                    {
                                        num27.Add(l);
                                        break;
                                    }
                                    if (l == 28)
                                    {
                                        num28.Add(l);
                                        break;
                                    }
                                    if (l == 29)
                                    {
                                        num29.Add(l);
                                        break;
                                    }
                                    if (l == 30)
                                    {
                                        num30.Add(l);
                                        break;
                                    }
                                    if (l == 31)
                                    {
                                        num31.Add(l);
                                        break;
                                    }
                                    if (l == 32)
                                    {
                                        num32.Add(l);
                                        break;
                                    }
                                    if (l == 33)
                                    {
                                        num33.Add(l);
                                        break;
                                    }
                                    if (l == 34)
                                    {
                                        num34.Add(l);
                                        break;
                                    }
                                    if (l == 35)
                                    {
                                        num35.Add(l);
                                        break;
                                    }
                                    if (l == 36)
                                    {
                                        num36.Add(l);
                                        break;
                                    }
                                    if (l == 37)
                                    {
                                        num37.Add(l);
                                        break;
                                    }
                                    if (l == 38)
                                    {
                                        num38.Add(l);
                                        break;
                                    }
                                    if (l == 39)
                                    {
                                        num39.Add(l);
                                        break;
                                    }
                                    if (l == 40)
                                    {
                                        num40.Add(l);
                                        break;
                                    }
                                    if (l == 41)
                                    {
                                        num41.Add(l);
                                        break;
                                    }
                                    if (l == 42)
                                    {
                                        num42.Add(l);
                                        break;
                                    }
                                    if (l == 43)
                                    {
                                        num43.Add(l);
                                        break;
                                    }
                                    if (l == 44)
                                    {
                                        num44.Add(l);
                                        break;
                                    }
                                    if (l == 45)
                                    {
                                        num45.Add(l);
                                        break;
                                    }
                                    if (l == 46)
                                    {
                                        num46.Add(l);
                                        break;
                                    }
                                    if (l == 47)
                                    {
                                        num47.Add(l);
                                        break;
                                    }
                                    if (l == 48)
                                    {
                                        num48.Add(l);
                                        break;
                                    }
                                    if (l == 49)
                                    {
                                        num49.Add(l);
                                        break;
                                    }
                                    if (l == 50)
                                    {
                                        num50.Add(l);
                                        break;
                                    }
                                    if (l == 51)
                                    {
                                        num51.Add(l);
                                        break;
                                    }
                                    if (l == 52)
                                    {
                                        num52.Add(l);
                                        break;
                                    }
                                    if (l == 53)
                                    {
                                        num53.Add(l);
                                        break;
                                    }
                                    if (l == 54)
                                    {
                                        num54.Add(l);
                                        break;
                                    }
                                    if (l == 55)
                                    {
                                        num55.Add(l);
                                        break;
                                    }
                                    if (l == 56)
                                    {
                                        num56.Add(l);
                                        break;
                                    }
                                    if (l == 57)
                                    {
                                        num57.Add(l);

                                        break;
                                    }
                                    if (l == 58)
                                    {
                                        num58.Add(l);
                                       
                                        break;
                                    }
                                    if (l == 59)
                                    {
                                        num59.Add(l);
                                        
                                        break;
                                    }
                                    if (l == 60)
                                    {
                                        num60.Add(l);
                                        
                                        break;
                                    }

                                }
                            }

                        }
                       
                        // Imprimir quebra de linha
                        Console.WriteLine(listGeral);
                    }
                    
                }
                
            }
            listGeral.Add(num1.Count);
            listGeral.Add(num2.Count);
            listGeral.Add(num3.Count);
            listGeral.Add(num4.Count);
            listGeral.Add(num5.Count);
            listGeral.Add(num6.Count);
            listGeral.Add(num7.Count);
            listGeral.Add(num8.Count);
            listGeral.Add(num9.Count);
            listGeral.Add(num10.Count);
            listGeral.Add(num11.Count);
            listGeral.Add(num12.Count);
            listGeral.Add(num13.Count);
            listGeral.Add(num14.Count);
            listGeral.Add(num15.Count);
            listGeral.Add(num16.Count);
            listGeral.Add(num17.Count);
            listGeral.Add(num18.Count);
            listGeral.Add(num19.Count);
            listGeral.Add(num20.Count);
            listGeral.Add(num21.Count);
            listGeral.Add(num22.Count);
            listGeral.Add(num23.Count);
            listGeral.Add(num24.Count);
            listGeral.Add(num25.Count);
            listGeral.Add(num26.Count);
            listGeral.Add(num27.Count);
            listGeral.Add(num28.Count);
            listGeral.Add(num29.Count);
            listGeral.Add(num30.Count);
            listGeral.Add(num31.Count);
            listGeral.Add(num32.Count);
            listGeral.Add(num33.Count);
            listGeral.Add(num34.Count);
            listGeral.Add(num35.Count);
            listGeral.Add(num36.Count);
            listGeral.Add(num37.Count);
            listGeral.Add(num38.Count);
            listGeral.Add(num39.Count);
            listGeral.Add(num40.Count);
            listGeral.Add(num41.Count);
            listGeral.Add(num42.Count);
            listGeral.Add(num43.Count);
            listGeral.Add(num44.Count);
            listGeral.Add(num45.Count);
            listGeral.Add(num46.Count);
            listGeral.Add(num47.Count);
            listGeral.Add(num48.Count);
            listGeral.Add(num49.Count);
            listGeral.Add(num50.Count);
            listGeral.Add(num51.Count);
            listGeral.Add(num52.Count);
            listGeral.Add(num53.Count);
            listGeral.Add(num54.Count);
            listGeral.Add(num55.Count);
            listGeral.Add(num56.Count);
            listGeral.Add(num57.Count);
            listGeral.Add(num58.Count);
            listGeral.Add(num59.Count);
            listGeral.Add(num60.Count);
            List<int> listGeralNova = new List<int>();
            foreach (var item in listGeral)
            {
                listGeralNova.Add(item);
            }
            maior = listGeralNova.Max();
            List<int> jogo = new List<int>();  
            for(int a = 0; a < 6; a++)
            {
                for (int u = 0; u < listGeral.Count; u++)
                {
                    if (maior == listGeral[u])
                    {
                        int num = u + 1;
                        jogo.Add(num);
                        for( int k =0; k<= listGeralNova.Count; k++)
                        {
                            if (maior == Convert.ToInt32(listGeralNova[k]))
                            {
                                listGeralNova.RemoveAt(k);
                                maior = listGeralNova.Max();
                                break;
                            }
                        }
                       
                    }
                    if(jogo.Count == 6)
                    {
                        break;
                    }
                    
                }
                if (jogo.Count == 6)
                {
                    break;
                }
            }
            Console.WriteLine("Seu numero de jogo são:");
            Console.WriteLine(""+jogo[0]+ ", " + jogo[1] + ", " + jogo[2] + ", " + jogo[3] + ", " + jogo[4] + ", " + jogo[5] + "");
            
        }

    }
}
