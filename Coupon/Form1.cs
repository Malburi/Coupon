using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;
using Newtonsoft.Json.Linq;

namespace Coupon
{
    public partial class Form1 : Form
    {
        private const string ApiUrl = "https://P11124-game-adapter.qookkagames.com/cms/active_code/change";

        private readonly List<string> playerNames = new List<string>();
        private readonly List<string> playerIds = new List<string>();
        private readonly List<string> ErrorPlayer = new List<string>();
        private readonly List<string> ErrorReason = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_ClickAsync(object sender, EventArgs e)
        {
            try
            {
                playerNames.Clear();
                playerIds.Clear();
                ErrorPlayer.Clear();
                ErrorReason.Clear();
                

                // Read data from Excel file (assuming you have a button to upload the Excel file)
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "Excel files|*.xls;*.xlsx";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    using (var stream = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            while (reader.Read())
                            {
                                // Assuming the first column contains player names and the second column contains player IDs
                                string playerName = reader.GetString(0);
                                string playerId = reader.GetString(1);

                                playerNames.Add(playerName);
                                playerIds.Add(playerId);
                            }
                        }
                    }
                }

                // Create an HttpClient to send the data
                using (var client = new HttpClient())
                {
                    for (int i = 0; i < playerNames.Count; i++)
                    {
                        var requestData = new
                        {
                            player_name = playerNames[i],
                            player_id = playerIds[i],
                            code = txt_coupon.Text // Replace with your actual code
                        };

                        var json = Newtonsoft.Json.JsonConvert.SerializeObject(requestData);
                        var content = new StringContent(json, Encoding.UTF8, "application/json");

                        var response = await client.PostAsync(ApiUrl, content);
                        var content2 = await response.Content.ReadAsStringAsync();
                        JObject parsedData = JObject.Parse(content2);

                        string response_message = parsedData["message"].ToString();
                        int code = (int)parsedData["code"];

                        switch (code)
                        {
                            case 200:
                                response_message = "교환 성공!";
                                break;
                            case 419:
                                response_message = "해당 쿠폰코드는 최대 교환 인원수를 초과하였거나 존재하지 않는 쿠폰코드입니다.";
                                break;
                            case 10608:
                                response_message = "잘못된 캐릭터 ID 혹은 캐릭터명입니다. 다시 입력해 주세요.";
                                break;
                            case 10610:
                                response_message = "귀하는 이미 해당 쿠폰코드와 중복 사용 불가한 다른 쿠폰코드를 사용했습니다.";
                                break;
                            case 10612:
                                response_message = "귀하는 이미 해당 쿠폰코드를 교환하여 중복 교환이 불가합니다!";
                                break;
                        }
                        if(playerNames[i] != null)
                        {
                            ErrorReason.Add(playerNames[i].ToString() + " : " + response_message);
                        }
                    }

                    // Create a multiline message to display the names
                    string message = string.Join(Environment.NewLine, ErrorReason);

                    // Show the message in a popup
                    MessageBox.Show(message, "Name List", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류 발생: {ex.Message}");
            }
        }
    }
}

