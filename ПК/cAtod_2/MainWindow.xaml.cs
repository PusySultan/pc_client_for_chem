using System;
using System.Collections;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using InTheHand.Net;
using InTheHand.Net.Bluetooth;
using InTheHand.Net.Sockets;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace cAtod_2
{
    public partial class MainWindow : Window
    {
        BluetoothClient bluetoothClient;
        bool bluetoothError = false;
        public MainWindow()
        {
            InitializeComponent();

            ChartUIP.Plot.Axes.Bottom.Label.Text = "Time";  // Настройка подписи нижней оси
            ChartUIP.Plot.Axes.Left.Label.Text = "Amplitude"; // Настройка подписи левой оси
            // ChartUIP.Plot.FigureBackground.Color = ScottPlot.Color.FromARGB(0xFF9B9696); // Настройка цвета графика#CFD8DC
            ChartUIP.Plot.FigureBackground.Color = ScottPlot.Color.FromHex("#CFD8DC");

            ChartPWM.Plot.Axes.Bottom.Label.Text = "Time";  // Настройка подписи гижней оси
            ChartPWM.Plot.Axes.Left.Label.Text = "Duty"; // Настройка подписи левой оси
            //ChartPWM.Plot.FigureBackground.Color = ScottPlot.Color.FromARGB(0xFF9B9696); // Настройка цвета графика
            ChartPWM.Plot.FigureBackground.Color = ScottPlot.Color.FromHex("#CFD8DC");

            ChartReact.Plot.Axes.Bottom.Label.Text = "Time";  // Настройка подписи нижней оси
            ChartReact.Plot.Axes.Left.Label.Text = "Pressure"; // Настройка подписи левой оси
            ChartReact.Plot.Axes.Right.Label.Text = "Temp"; // Настройка подписи правой оси
            //ChartReact.Plot.FigureBackground.Color = ScottPlot.Color.FromARGB(0xFF9B9696); // Настройка цвета графика
            ChartReact.Plot.FigureBackground.Color = ScottPlot.Color.FromHex("#CFD8DC");

            ChartVodorod.Plot.Axes.Bottom.Label.Text = "Time";  // Настройка подписи нижней оси
            ChartVodorod.Plot.Axes.Left.Label.Text = "Pressure"; // Настройка подписи левой оси
            ChartVodorod.Plot.Axes.Right.Label.Text = "Temp"; // Настройка подписи правой оси
            //ChartVodorod.Plot.FigureBackground.Color = ScottPlot.Color.FromARGB(0xFF9B9696); // Настройка цвета графика
            ChartVodorod.Plot.FigureBackground.Color = ScottPlot.Color.FromHex("#CFD8DC");

            ImpulseVisualisation.Plot.Axes.Bottom.Label.Text = "Time";  // Настройка подписи гижней оси
            ImpulseVisualisation.Plot.Axes.Left.Label.Text = "A"; // Настройка подписи левой оси
            //ChartPWM.Plot.FigureBackground.Color = ScottPlot.Color.FromARGB(0xFF9B9696); // Настройка цвета графика
            ImpulseVisualisation.Plot.FigureBackground.Color = ScottPlot.Color.FromHex("#CFD8DC");

            LineVisualisationMode.Plot.Axes.Bottom.Label.Text = "Time";
            LineVisualisationMode.Plot.Axes.Left.Label.Text = "A";
            LineVisualisationMode.Plot.FigureBackground.Color = ScottPlot.Color.FromHex("#CFD8DC");

            try
            {
                bluetoothClient = new BluetoothClient();
            }
            catch (Exception ex)
            {
                bluetoothError = true;
                processTXT.Text = "Ошибка, возможно bluetooth на устройстве выключен";
            }

        }

        ArrayList devise = new ArrayList();
        bool isRecevedData = true;

        List<double> dataUIPy = new List<double>(); // Динамический массив для оси У UIP  

        List<double> dataPWMy = new List<double>(); // Динамический массив для оси У PWM  

        List<double> dataReactTy = new List<double>(); // Динамический массив для оси У ReactT

        List<double> dataReactDy = new List<double>(); // Динамический массив для оси У ReactD

        List<double> dataVodorodTy = new List<double>(); // Динамический массив для оси У VodorodT  

        List<double> dataVodorodDy = new List<double>(); // Динамический массив для оси У VodorodD  

        string cpuTemp = "";
        string sostKulera = "";
        string status = "";

        int numChoseDevise = -1;


        // Поиск устройств
        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            if (bluetoothError)
            {
                try
                {
                    bluetoothClient = new BluetoothClient();
                }
                catch (Exception ex)
                {
                    bluetoothError = true;
                    processTXT.Text = ex.Message;
                    return;
                }

                bluetoothError = false;
            }

            listB.Items.Clear();
            devise.Clear();

            processTXT.Text = "Выполняется поиск устройств...";
            await serchBLE_async();

        }

        // Определение асинхронной функции поиска устройств
        async Task serchBLE_async()
        {
            BluetoothDeviceInfo[] mass = null;

            try
            {
                // Определяем что данные в массив нужно ждать 
                mass = await Task.Run(() => bluetoothClient.DiscoverDevices());
            }
            catch (Exception ex)
            {
                processTXT.Text = ex.Message;
                bluetoothError = true;
                return;
            }

            for (int i = 0; i < mass.Length; i++)
            {

                devise.Add(
                    new Device(
                         mass[i].DeviceName,
                         new BluetoothAddress(mass[i].DeviceAddress.ToByteArray()),
                         new BluetoothEndPoint(new BluetoothAddress(mass[i].DeviceAddress.ToByteArray()), BluetoothService.SerialPort)
              ));

            }

            for (int i = 0; i < devise.Count; i++)
            {
                var item = devise[i] as Device;
                listB.Items.Add(item.blName);
            }

            processTXT.Text = "Сканирование устройств завершено";
        }

        // Выбор устройства из списка
        private async void listB_SelectionChanged(object sender, SelectionChangedEventArgs e) // SelectionChanged="listB_SelectionChanged 
        {

            if (numChoseDevise == -1)
            {
                BluetoothEndPoint blepItem = null;
                string blNAme = "";

                for (int i = 0; i < devise.Count; i++)
                {
                    var item = devise[i] as Device;
                    if (listB.SelectedItem.ToString() == item.blName)
                    {
                        blepItem = item.blEndpoint;
                        blNAme = item.blName;
                        processTXT.Text = "Выполняется подключение";
                        numChoseDevise = i;
                        break;
                    }

                }

                if (blepItem != null)
                {
                    await connectBLE_async(blepItem, blNAme); // THIS await delete

                }
                else
                {
                    processTXT.Text = "Ошибка";
                }
            }

        }

        async Task connectBLE_async(BluetoothEndPoint blEndPoint, String blName) // Определение асинхронной функции подключения к устройству
        {
            await Task.Run(() => {

                try
                {
                    bluetoothClient.Connect(blEndPoint);
                }
                catch (Exception ex)
                {
                    // processTXT.Text = ex.Message;
                    bluetoothClient.Close();
                }

            });

            if (bluetoothClient.Connected)
            {
                processTXT.Text = "Подключено к " + blName;
            }
            else {
                processTXT.Text = "Не подключено";
                bluetoothClient.Close();
            }
        }

        // Получение данных
        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (bluetoothError)
            { 
                return;
            }

            if (bluetoothClient.Connected)
            {

                string dataToSend = "1"; // Команда для отправки
                string ReceveData = "";  // Принятые данные

                NetworkStream stream = bluetoothClient.GetStream(); // Получаем входной поток
                byte[] receiveBuffer = new byte[1024];    // Буфер для записи принятых данных
                byte[] dataBytestoSend = Encoding.UTF8.GetBytes(dataToSend); // массив байтов для отпраки команды на МК

                try
                {
                    while (isRecevedData)
                    {
                        stream.Write(dataBytestoSend, 0, dataBytestoSend.Length);// Отправили команду старта приема на МК

                        processTXT.Text = "Идет процесс получения даных";
                        ReceveData = await recevedData_async(stream, receiveBuffer);

                        try
                        {
                            dynamic json = JsonConvert.DeserializeObject(ReceveData);

                            if (Check_UIP.IsChecked == false) //UIP
                            {
                                ChartUIP.Plot.Clear();
                                dataUIPy.Add(Convert.ToDouble(json.OutP));

                                ChartUIP.Plot.Add.Signal(dataUIPy);
                                ChartUIP.Refresh();
                                ChartUIP.Plot.Axes.AutoScale();
                            }

                            if (Check_PWM.IsChecked == false)
                            {
                                ChartPWM.Plot.Clear();
                                dataPWMy.Add(Convert.ToDouble(json.PWM));

                                ChartPWM.Plot.Add.Signal(dataPWMy);
                                ChartPWM.Plot.Axes.AutoScale(); // Автоматический маштаб
                                ChartPWM.Refresh();
                            }

                            if (Check_React.IsChecked == false)
                            {
                                ChartReact.Plot.Clear();
                                // Находим температуру реактора
                                dataReactTy.Add(Convert.ToDouble(json.ReactT));

                                // Находим давление в реакторе
                                dataReactDy.Add(Convert.ToDouble(json.ReactD));

                                // Добавляем температуру реактора
                                var ReactTemp = ChartReact.Plot.Add.Signal(dataReactTy);

                                // Добавляем давление в реакторе
                                var ReactDuit = ChartReact.Plot.Add.Signal(dataReactDy);


                                ReactTemp.LegendText = "Т реактора";
                                ReactDuit.LegendText = "Д реактора";

                                ChartReact.Plot.Legend.IsVisible = true;
                                ChartReact.Plot.Legend.Alignment = ScottPlot.Alignment.LowerRight;

                                ChartReact.Plot.Axes.AutoScale();// Автоматический маштаб
                                ChartReact.Refresh();

                                //ChartReact.Plot.Legend
                            }

                            if (Check_Vodorod.IsChecked == false)
                            {
                                ChartVodorod.Plot.Clear();

                                // Находим температуру водородного элемента
                                dataVodorodTy.Add(Convert.ToDouble(json.Vodorod_T));

                                // Добавляем температуру водородного элемента
                                var VodorodT = ChartVodorod.Plot.Add.Signal(dataVodorodTy);

                                // Находим давление в водородном элементе
                                dataVodorodDy.Add(Convert.ToDouble(json.Vodorod_D));

                                // Добавляем давление в водородном элементе
                                var VodododD = ChartVodorod.Plot.Add.Signal(dataVodorodDy);

                                VodorodT.LegendText = "Т вод.эл";
                                VodododD.LegendText = "Д вод.эл";

                                ChartVodorod.Plot.Legend.IsVisible = true;
                                ChartVodorod.Plot.Legend.Alignment = ScottPlot.Alignment.LowerRight;

                                ChartVodorod.Refresh();
                                ChartVodorod.Plot.Axes.AutoScale();
                            }


                            CPU_temp.Text = json.CPU_temp;
                            Kuler_status.Text = json.Sost_kul;
                            Status.Text = json.Status;
                        }
                        catch (Exception ex) { }
                    }

                    stream.Close();
                    bluetoothClient.Close();

                    processTXT.Text = "Устройств Bluetooth отключено";

                }
                catch (Exception ex)
                {
                    processTXT.Text = ex.Message;
                    stream.Close();
                    bluetoothClient.Close();
                }
            }
        }

        // Определение асинхронной функции для приема данных
        async Task<string> recevedData_async(NetworkStream stream, byte[] receiveBuffer)
        {
            int bytesRead = 0;

            await Task.Run(() => {

                bytesRead = stream.Read(receiveBuffer, 0, receiveBuffer.Length);// возвращает колличество сичитанных байт

            });

            return Encoding.UTF8.GetString(receiveBuffer, 0, bytesRead);
        }


        // Отключиться от устройства 
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (isRecevedData)
            {
                isRecevedData = !isRecevedData; // В асинхронной функции получения данных происходит отключение от устройства
            }
        }

        private void expEXCEL_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application application = null; // Создание объекта документа

            Excel.Workbooks workbooks = null;  // Создание объекта книг
            Excel.Workbook workbook = null;  // Создание объекта книги

            Excel.Sheets worksheets = null;  // Создание объекта страниц
            Excel.Worksheet worksheet = null;  // Создание объекта страницы

            Excel.Range cell = null; // Создание обьекта ячейки

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "EXCEL_cAtod files (*.xlsx)|*.xlsx| csv_cAtod files (*.csv)|*.csv";

            try
            {
                application = new Excel.Application();//Инициализация дока
                workbooks = application.Workbooks;//Инициализация книг
                workbook = workbooks.Add(); // Инициализация книги
                worksheets = workbook.Worksheets; //получаем доступ к коллекции рабочих листов 
                worksheet = worksheets.Item[1]; //получаем доступ к первому листу

                if (onUIPch.IsChecked == true)
                {
                    cell = worksheet.Cells[1, 1];
                    cell.Value = "Выходная мощность";
                    for (int i = 0; i < dataUIPy.Count; i++)
                    {
                        cell = worksheet.Cells[i + 2, 1];//получаем доступ к ячейке
                        cell.Value = dataUIPy[i];
                    }

                    cell = worksheet.Cells[1, 2];
                    cell.Value = "Время, с";
                    for (int i = 0; i < dataUIPy.Count; i++)
                    {
                        cell = worksheet.Cells[i + 2, 2];
                        cell.Value = i;
                    }
                }



                if (onPWMch.IsChecked == true)
                {
                    cell = worksheet.Cells[1, 4];
                    cell.Value = "Скважность ШИМ";
                    for (int i = 0; i < dataPWMy.Count; i++)
                    {
                        cell = worksheet.Cells[i + 2, 4];//получаем доступ к ячейке
                        cell.Value = dataPWMy[i];
                    }

                    cell = worksheet.Cells[1, 5];
                    cell.Value = "Время, с";
                    for (int i = 0; i < dataPWMy.Count; i++)
                    {
                        cell = worksheet.Cells[i + 2, 5];
                        cell.Value = i;
                    }
                }




                if (onDuitReactch.IsChecked == true || onTempReactch.IsChecked == true)
                {
                    if (onTempReactch.IsChecked == true)
                    {
                        cell = worksheet.Cells[1, 7];
                        cell.Value = "Температура в реакторе";
                        for (int i = 0; i < dataReactTy.Count; i++)
                        {
                            cell = worksheet.Cells[i + 2, 7];//получаем доступ к ячейке
                            cell.Value = dataReactTy[i];
                        }
                    }

                    if (onDuitReactch.IsChecked == true)
                    {
                        cell = worksheet.Cells[1, 8];
                        cell.Value = "Давление в реакторе";
                        for (int i = 0; i < dataReactDy.Count; i++)
                        {
                            cell = worksheet.Cells[i + 2, 8];//получаем доступ к ячейке
                            cell.Value = dataReactDy[i];
                        }
                    }

                    cell = worksheet.Cells[1, 9];
                    cell.Value = "Время, с";
                    for (int i = 0; i < dataReactDy.Count; i++)
                    {
                        cell = worksheet.Cells[i + 2, 9];
                        cell.Value = i;
                    }
                }



                if (onDuitVodorodch.IsChecked == true || onTempVodorodch.IsChecked == true)
                {

                    if (onTempVodorodch.IsChecked == true)
                    {
                        cell = worksheet.Cells[1, 11];
                        cell.Value = "Температура в водородном элементе";
                        for (int i = 0; i < dataVodorodTy.Count; i++)
                        {
                            cell = worksheet.Cells[i + 2, 11];//получаем доступ к ячейке
                            cell.Value = dataVodorodTy[i];
                        }
                    }

                    if (onDuitVodorodch.IsChecked == true)
                    {
                        cell = worksheet.Cells[1, 12];
                        cell.Value = "Давление в водородном элементе";
                        for (int i = 0; i < dataVodorodDy.Count; i++)
                        {
                            cell = worksheet.Cells[i + 2, 12];//получаем доступ к ячейке
                            cell.Value = dataVodorodDy[i];
                        }
                    }

                    cell = worksheet.Cells[1, 13];
                    cell.Value = "Время, с";
                    for (int i = 0; i < dataReactDy.Count; i++)
                    {
                        cell = worksheet.Cells[i + 2, 13];
                        cell.Value = i;
                    }

                }




                if (saveFileDialog.ShowDialog() != DialogResult)
                {
                    workbook.SaveAs(saveFileDialog.FileName); // выбор места сохранения
                    application.Quit();//Закрываем приложение
                }
                else
                {
                    application.Quit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "");
            }
            finally {
                Marshal.ReleaseComObject(cell);//Удаляем из оперативы
                Marshal.ReleaseComObject(worksheet);//Удаляем из оперативы
                Marshal.ReleaseComObject(worksheets);//Удаляем из оперативы
                Marshal.ReleaseComObject(workbook);//Удаляем из оперативы
                Marshal.ReleaseComObject(workbooks);//Удаляем из оперативы
                Marshal.ReleaseComObject(application);//Удаляем из оперативы
            }


        }

        private void clearPlot_but_Click(object sender, RoutedEventArgs e)
        {
            dataUIPy.Clear();
            ChartUIP.Plot.Clear();
            ChartUIP.Refresh();

            dataPWMy.Clear();
            ChartPWM.Plot.Clear();
            ChartPWM.Refresh();

            dataReactTy.Clear();
            dataReactDy.Clear();
            ChartReact.Plot.Clear();
            ChartReact.Refresh();

            dataVodorodDy.Clear();
            dataVodorodTy.Clear();
            ChartVodorod.Plot.Clear();
            ChartVodorod.Refresh();
        }

        private void previewBut_Click(object sender, RoutedEventArgs e)
        {
            double borehole = 0;
            double[] PWMx = new double[] { 1, 1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 5, 7, 7, 7, 7, 9, 9, 9, 9 };
            double[] PWMy = new double[] { 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0 };

            byte[] cPoint = new byte[] { 2, 3, 6, 7, 10, 11, 14, 15 };

            try
            {
                borehole = double.Parse(BoreholeInput.Text);

                if (borehole > 99.999 || borehole < 0.001)
                {
                    MessageBox.Show("Скважность не может быть больше 99.999% и меньше 0.001%", "Внимание");
                    return;
                }

                foreach (var item in cPoint)
                {
                    PWMx[item] += 2 * (borehole / 100);
                }

                LineVisualisationMode.Plot.Clear();
                LineVisualisationMode.Plot.Add.Scatter(PWMx, PWMy);
                LineVisualisationMode.Plot.Axes.AutoScale();// Автоматический маштаб
                LineVisualisationMode.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Внимание");
            }
        }

        private void SendCommand_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                NetworkStream stream = bluetoothClient.GetStream(); // Получаем входной поток
                Parsel parsel = new Parsel();

                if(tabeMode.SelectedIndex == 0) 
                { 
                    parsel.borehole = double.Parse(boreholeInput.Text);
                    parsel.jobTime = double.Parse(jobTimeInput.Text);
                    parsel.pauseTime = double.Parse(pauseTimeInput.Text);
                    parsel.mode = "Pulse";
                
                }
                else
                {
                    parsel.borehole = double.Parse(BoreholeInput.Text);
                    parsel.jobTime = null;
                    parsel.pauseTime = null;
                    parsel.mode = "Line";

                }
               
                string outData = JsonConvert.SerializeObject(parsel); // Превращаем обьект в  JSON
                byte[] dataBytestoSend = Encoding.UTF8.GetBytes(outData);

                stream.Write(dataBytestoSend, 0, dataBytestoSend.Length);// Отправили команду старта приема на МК

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка");
                // parsel.borehole = 0.0;
            }
        }

        private void previewImpulse_Click(object sender, RoutedEventArgs e)
        {
            List<double> x = new List<double>();
            List<double> y = new List<double>();

            double[] PWMx = new double[] { 1, 1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 5, 7, 7, 7, 7, 9, 9, 9, 9, 11, 11, 11, 11 };
            double[] PWMy = new double[] { 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0 ,1 ,1, 0, 0,   1,  1,  0 };
            byte[] cPoint = new byte[]   { 2, 3, 6, 7, 10, 11, 14, 15 };

            double borehole = 50;
            double jobtime = 10;
            double pauseTime = 5;

            try
            {
                borehole = double.Parse(boreholeInput.Text);
                jobtime = double.Parse(jobTimeInput.Text);
                pauseTime = double.Parse(pauseTimeInput.Text);

                if (borehole > 99.999 || borehole < 0.001)
                {
                    MessageBox.Show("Скважность не может быть больше 99.999% и меньше 0.001%", "Внимание");
                    return;
                }

                foreach (var item in cPoint) // Установка скважности
                {
                    PWMx[item] += 2 * (borehole / 100);
                }

                int? index = ClearArrX(PWMx, jobtime);
                x.AddRange(PWMx);
                y.AddRange(PWMy);
                
                if (index != null)
                {
                    x.RemoveRange((int)index, PWMx.Length - (int)index);  
                    y.RemoveRange((int)index, PWMy.Length - (int)index);

                    x.Add(x[(int)index - 1]);
                    x.Add(jobtime + pauseTime);
                    y.Add(0); y.Add(0);

                    for (int i = 0; i < (int) index; i++)
                    {
                        x.Add( (x[i] + (jobtime + pauseTime)) );
                        y.Add(y[i]);
                    };
                }

               

                ImpulseVisualisation.Plot.Clear();
                ImpulseVisualisation.Plot.Add.Scatter(x, y);
                ImpulseVisualisation.Plot.Axes.AutoScale();// Автоматический маштаб
                ImpulseVisualisation.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Внимание");
            }
        }

        private int? ClearArrX(double[]Arr, double jobTime)
        {
            for (int i = 0; i < Arr.Length; i++)
            {
                if (Arr[i] > jobTime)
                {
                    return i;
                }
            }

            return null;
        }
    }
}
