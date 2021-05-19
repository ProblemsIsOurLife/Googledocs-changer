using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.Threading;


namespace GDLPSC
{
   
    public partial class Form1 : Form
    {
        bool inwork = false;
        private static string[,] data = new string[10, 3];
        private static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private static readonly string AppName = "DBLPC";
        //Id дока
        private static readonly string SpreadsheetId = "1la6IP9FReLgd_Xf-QIdMaeie7_LCvuEjAPMKYocpdSE";
        private String Range = "A1:C10";
        public Form1()
        {
            
            InitializeComponent();
            
        }

        
        private void Button1_Click(object sender, EventArgs e)
        {
            inwork = true;

            dataGridView1.MultiSelect = false;

            dataGridView1.RowCount = 10;
            dataGridView1.ColumnCount = 3;
            for (int i = 0; i < 10; i++)
                for (int j = 0; j < 3; j++)
                    data[i, j]="";
            var credential = GetSheetCredentionals();
           var service = GetService(credential);
           GetCells(service, Range, SpreadsheetId);
            for(int i = 0; i < 10; i++)  
                  for(int j=0; j<3;j++)
                 dataGridView1.Rows[i].Cells[j].Value = data[i,j];
            for (int j = 0; j < 3; j++) {
                dataGridView1.Columns[j].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            pictureBox1.Visible = Visible;

        }


        private void Button2_Click(object sender, EventArgs e)
        {
            if (inwork == true)
            {
                for (int i = 0; i < 10; i++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        data[i, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }

                }
                string message2 = "Данные успешно сохранены";
                string caption2 = "Успех";
                MessageBoxButtons buttons2 = MessageBoxButtons.OK;
                DialogResult result2;
                inwork = true;
                result2 = MessageBox.Show(message2, caption2, buttons2);
               
               }
            else
            {
                string message = "Нет данных. Загрузить?";
                string caption = "Error";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result;
                inwork = true;
                result = MessageBox.Show(message, caption, buttons);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    dataGridView1.RowCount = 10;
                    dataGridView1.ColumnCount = 3;
                    var credential1 = GetSheetCredentionals();
                    var service1 = GetService(credential1);
                    GetCells(service1, Range, SpreadsheetId);
                    for (int i = 0; i < 10; i++)
                        for (int j = 0; j < 3; j++)
                            dataGridView1.Rows[i].Cells[j].Value = data[i, j];
                    for (int j = 0; j < 3; j++)
                    {
                        dataGridView1.Columns[j].SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }
                pictureBox1.Visible = Visible;
            }





            var credential = GetSheetCredentionals();
            var service = GetService(credential);
            FillSpreadsheet(service, SpreadsheetId, data);
        }




        private static UserCredential GetSheetCredentionals()
        {  
            //файл авторизации
            using (var stream =
               new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }
        }
        private static SheetsService GetService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = AppName
            });
        }
        private static void FillSpreadsheet(SheetsService service, string SpreadsheetId, string[,] data)
        {
            List<Request> requests = new List<Request>();
            for (int i=0; i < data.GetLength(0); i++)
            {
                List<CellData> values = new List<CellData>();
                for (int j = 0;j < data.GetLength(1); j++)
                {
                    values.Add(new CellData
                    {
                        UserEnteredValue = new ExtendedValue
                        {
                            StringValue = data[i, j]
                        }
                    });
                }
                requests.Add(
                    new Request
                    {
                        UpdateCells = new UpdateCellsRequest
                        {
                            Start = new GridCoordinate
                            {
                                SheetId = 0,
                                RowIndex = i,
                                ColumnIndex = 0
                            },
                            Rows = new List<RowData> { new RowData { Values = values } },
                            Fields = "userEnteredValue"
                        }
                    }
                    ); 
            }
            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };
            service.Spreadsheets.BatchUpdate(busr, SpreadsheetId).Execute();
        }
        
        private static void GetCells(SheetsService service, string range, string spreadsheetId)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
            
            int i = 0;
            foreach (var value in response.Values)
            {
                data[i, 0] += value[0];
                data[i, 1] += value[1];
                data[i, 2] += value[2];
                i++;
            }
            
        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {
                DataGridViewCell cell = dataGridView1.CurrentCell;
                dataGridView1.CurrentCell = cell;
                dataGridView1.BeginEdit(true);         
        }
    }
    
}
