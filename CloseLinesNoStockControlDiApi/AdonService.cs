using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace CloseLinesNoStockControlDiApi
{
    public class AdonService
    {
        public Company Company { get; private set; }

        public string Server { get; set; }
        public string CompanyDB { get; set; }
        public string DbUserName { get; set; }
        public string DbPassword { get; set; }
        public BoDataServerTypes DbServerType { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }

        public AdonService()
        {
            Company = new Company();
        }

        public void ConnectSbo()
        {
            Company.Server = Server;
            Company.CompanyDB = CompanyDB;
            Company.DbUserName = DbUserName;
            Company.DbPassword = DbPassword;
            Company.DbServerType = BoDataServerTypes.dst_MSSQL2014;
            Company.UserName = UserName;
            Company.Password = Password;

            Company.UseTrusted = false;
            Company.language = BoSuppLangs.ln_Portuguese_Br;
            Company.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;

            if (Company.Connect() != 0)
                throw new Exception("Erro ao conectar no sap:" +
                           $"[{Company.GetLastErrorCode()}] - [{Company.GetLastErrorDescription()}]");
        }

        private string GetConnectionString()
        {
            var connStr = new SqlConnectionStringBuilder();
            connStr.DataSource = Server;
            connStr.InitialCatalog = CompanyDB;
            connStr.UserID = DbUserName;
            connStr.Password = DbPassword;
            return connStr.ToString();
        }

        public void CloseLinesOfQuotation(int docEntry)
        {
            var businessObject = Company.GetBusinessObject(BoObjectTypes.oQuotations) as Documents;

            if (!businessObject.GetByKey(docEntry))
                throw new Exception($"Não foi possível localizar a cotação N°[{docEntry}] no SAP.");

            var itensToCloseLines = GeItensWithNotStockControl(businessObject);
            SetContractId(businessObject);
            CloseLines(businessObject, itensToCloseLines);
            UpdateDocument(businessObject);
        }

        private void CloseLines(Documents businessObject, IList<string> itensToCloseLines)
        {
            for (var i = 0; i < businessObject.Lines.Count; i++)
            {
                businessObject.Lines.SetCurrentLine(i);

                if (itensToCloseLines.Any(x => x == businessObject.Lines.ItemCode) &&
                    businessObject.Lines.LineStatus == BoStatus.bost_Open)
                    businessObject.Lines.LineStatus = BoStatus.bost_Close;
            }
        }

        private void UpdateDocument(Documents businessObject)
        {
            if (businessObject.Update() != 0)
                throw new Exception($"Erro ao atualizar documento no SAP.\n[{Company.GetLastErrorCode()}]-[{Company.GetLastErrorDescription()}]");
        }

        private void SetContractId(Documents businessObject)
        {
            businessObject.UserFields.Fields.Item("U_ContractId").Value = 1;
            UpdateDocument(businessObject);
        }

        private IList<string> GeItensWithNotStockControl(Documents businessObject)
        {
            var itemsCode = new List<string>();

            for (var i = 0; i < businessObject.Lines.Count; i++)
            {
                businessObject.Lines.SetCurrentLine(i);
                itemsCode.Add(businessObject.Lines.ItemCode);
            }

            return GeItensWithNotStockControlFromDb(itemsCode);
        }

        private IList<string> GeItensWithNotStockControlFromDb(IList<string> listItemCode)
        {
            var itemsCode = new List<string>();
            var data = GetItemCodeData(listItemCode);

            foreach (DataRow row in data.Rows)
            {
                var itemCode = row.Field<string>("ItemCode");
                itemsCode.Add(itemCode);
            }

            return itemsCode;
        }

        private DataTable GetItemCodeData(IList<string> listItemCode)
        {
            var connectionString = GetConnectionString();
            var data = new DataTable();

            using (var conn = new SqlConnection(connectionString))
            {
                var command = conn.CreateCommand();
                command.CommandType = CommandType.Text;
                command.CommandText = GetQuery(listItemCode);

                conn.Open();
                if (conn.State != ConnectionState.Open)
                    throw new Exception("Não foi possível conectar ao banco de dados");

                var reader = command.ExecuteReader();
                data.Load(reader);
                conn.Close();
            }

            return data;
        }

        private string GetQuery(IList<string> listItemCode)
        {
            var sb = new StringBuilder($@"SELECT [ItemCode] FROM [{CompanyDB}]..[OITM] 
                                         WHERE [InvntItem] = 'N'
                                         AND [ItemCode] IN(");

            foreach (var itemCode in listItemCode)
            {
                if (itemCode == listItemCode.Last())
                    sb.Append($"'{itemCode}');");
                else
                    sb.Append($"'{itemCode}',");
            }

            return sb.ToString();
        }
    }
}
