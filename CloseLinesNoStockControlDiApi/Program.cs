using SAPbobsCOM;
using System;

namespace CloseLinesNoStockControlDiApi
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Informe o código da cotação para realizar o fechamento das linhas!");
                var docEntry = Console.ReadLine();

                var service = new AdonService();
                SetParams(service);
                service.ConnectSbo();
                service.CloseLinesOfQuotation(Convert.ToInt32(docEntry));

                PrintSuccess();
            }
            catch (Exception ex)
            {
                PrintMessageError(ex.Message);
            }
        }

        private static void SetParams(AdonService service)
        {
            service.Server = @"sap91-pc\b1";
            service.CompanyDB = "SBO_AREA_SEM_ADDON";
            service.DbUserName = "sa";
            service.DbPassword = "sap@123";
            service.DbServerType = BoDataServerTypes.dst_MSSQL2014;
            service.UserName = "manager";
            service.Password = "sapbone";
        }

        private static void PrintSuccess()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine("Operação realizada com sucesso");
            Console.ReadKey();
        }

        private static void PrintMessageError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine(message);
            Console.ResetColor();
            Console.ReadKey();
        }
    }
}
