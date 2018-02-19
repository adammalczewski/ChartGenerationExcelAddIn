using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddIn1 {
    public partial class ThisAddIn {

        Office.CommandBarButton _drawGraphButton, _drawAllGraphsButton1, _drawAllGraphsButton2;

        const int _GRAPH_WIDTH = 8;
        const int _GRAPH_HEIGHT = 14;

        const int _COLUMNS_DISTANCE = 14;

        const int _MAX_SEARCH = 10;

        private string GetGraphName(Worksheet worksheet) {
            int num = 1;
            string name = "Graph" + 1;
            while (worksheet.Controls.Contains(name)) {
                ++num;
                name = "Graph" + num;
            }
            return name;
        }

        private void DrawGraphFromSelectedCells(int graphRow, int graphColumn,string title,string valueName, Worksheet worksheet) {
            bool screenUpdating = Application.ScreenUpdating;
            Application.ScreenUpdating = false;
            Excel.Application exApp = Globals.ThisAddIn.Application;
            Excel.Range selectedRange = exApp.Selection;
            long row = selectedRange.Row;
            long col = selectedRange.Column;
            
            Worksheet activeWorksheet = Globals.Factory.GetVstoObject(Application.ActiveWorkbook.ActiveSheet);

            string name = GetGraphName(activeWorksheet);

            Chart chart = activeWorksheet.Controls.AddChart(activeWorksheet.Range["A1"].Resize[_GRAPH_HEIGHT,_GRAPH_WIDTH]
                .Offset[graphRow*_GRAPH_HEIGHT,graphColumn*_GRAPH_WIDTH], name);

            chart.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;
            chart.SetSourceData(selectedRange, Excel.XlRowCol.xlColumns);
            chart.HasTitle = true;
            chart.ChartTitle.Text = title;
            Excel.Axis xAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
            xAxis.HasTitle = true;
            xAxis.AxisTitle.Text = activeWorksheet.Cells[row, col].Text;
            Excel.Axis yAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Text = valueName;
            xAxis.MinimumScaleIsAuto = false;
            Excel.Name newName = activeWorksheet.Names.Add("firstCol", selectedRange.Resize[selectedRange.Rows.Count, 1], false);
            xAxis.MinimumScale = (double)activeWorksheet.Evaluate("MIN(firstCol)");
            xAxis.MaximumScale = (double)activeWorksheet.Evaluate("MAX(firstCol)");
            newName.Delete();
            newName = activeWorksheet.Names.Add("col", selectedRange.Resize[selectedRange.Rows.Count, 1].Offset[missing,1], false);
            double maxValue = (double) activeWorksheet.Evaluate("MAX(col)");
            double minValue = (double) activeWorksheet.Evaluate("MIN(col)");
            newName.Delete();
            for (int i = 2; i < selectedRange.Columns.Count; ++i) {
                newName = activeWorksheet.Names.Add("col", selectedRange.Resize[selectedRange.Rows.Count, 1].Offset[missing, i], false);
                double maxValueNow = (double)activeWorksheet.Evaluate("MAX(col)");
                double minValueNow = (double)activeWorksheet.Evaluate("MIN(col)");
                newName.Delete();
                if (maxValueNow > maxValue) maxValue = maxValueNow;
                if (minValueNow < minValue) minValue = minValueNow;
            }
            yAxis.MinimumScaleIsAuto = false;
            yAxis.MinimumScale = minValue;
            yAxis.MaximumScale = maxValue;
            chart.Location(Excel.XlChartLocation.xlLocationAsObject,worksheet.Name);
            activeWorksheet.Activate();
            Application.ScreenUpdating = screenUpdating;

        }

        private void DrawGraph(Office.CommandBarButton cmdBarbutton, ref bool cancel) {

            GetGraphsAndTablesWorksheet();
            Worksheet worksheet = (Worksheet) Globals.Factory.GetVstoObject(Application.ActiveWorkbook.Sheets["Wykresy i Tabele"]);
            DrawGraphFromSelectedCells(0, 0, "Wykres", "", worksheet);
            

            DrawGraphFromSelectedCells(1, 1, "Wykres", "", worksheet);

            DrawGraphFromSelectedCells(4, 0, "Wykres", "", worksheet);

        }

        private class WorksheetReader {

            private readonly Worksheet _worksheet;

            public WorksheetReader(Worksheet worksheet) {
                _worksheet = worksheet;
            }

            public bool IsString(int row, int column) {
                return ((Excel.Range) _worksheet.Cells[row, column]).Value2 is string;
            }

            public string GetString(int row, int column) {
                return (_worksheet.Cells[row,column]).Value2;
            }

            public bool IsNumber(int row, int column) {
                return ((Excel.Range) _worksheet.Cells[row, column]).Value2 is double;
            }

            public int GetInt(int row, int column) {
                return (int) ((Excel.Range) _worksheet.Cells[row, column]).Value2;
            }

            public double GetDouble(int row, int column) {
                return (double)((Excel.Range)_worksheet.Cells[row, column]).Value2;
            }

        }

        private class Table {

            public readonly List<string> Parameters;
            public readonly List<string> Values;
            public string Name;

            public Table() {
                Parameters = new List<string>();
                Values = new List<string>();
            }

        }

        private class Data {

            public readonly Dictionary<string,int> Parameters;
            public readonly Dictionary<string, double> Values; 

            public Data() {
                Parameters = new Dictionary<string,int>();
                Values = new Dictionary<string, double>();
            }

        }


        private void DrawAllGraphs(int columns) {
            bool screenUpdating = Application.ScreenUpdating;
            Application.ScreenUpdating = false;
            Worksheet worksheet = Globals.Factory.GetVstoObject(this.Application.ActiveWorkbook.ActiveSheet);

            WorksheetReader reader = new WorksheetReader(worksheet);

            List<Table> tables = new List<Table>();
            List<Data> dataList = new List<Data>();

            GetGraphsAndTablesWorksheet();
            Worksheet newWorksheet = (Worksheet)Globals.Factory.GetVstoObject(Application.ActiveWorkbook.Sheets["Wykresy i Tabele"]);

            bool error = false;
            String errorString = "";

            List<String> parametersList = new List<String>();

            for (int c = 0; c < columns && !error; ++c) {

                bool foundSomething = true;
                int colTrans = 0;
                int graphNum = 0;

                while (foundSomething && !error) {

                    int column = 1 + (c + columns*colTrans)*_COLUMNS_DISTANCE;

                    bool foundNext = true;
                    int rowPos = 1;

                    for (int rowsSearched = 0; rowsSearched < _MAX_SEARCH && ((Excel.Range)worksheet.Cells[rowPos, column]).Value2 == null
                        ; ++rowsSearched) ++rowPos;

                    if (((Excel.Range) worksheet.Cells[rowPos, column]).Value2 == null) foundNext = false;

                    foundSomething = foundNext;

                    while (foundNext && !error) {

                        String blockType = null;

                        if (reader.IsString(rowPos, column)) blockType = reader.GetString(rowPos, column);

                        if (blockType == null || blockType[0] != '[' || blockType[blockType.Length-1] != ']') {
                            foundNext = false;
                            continue;
                        }

                        blockType = blockType.Substring(1, blockType.Length - 2);

                        switch (blockType) {
                            case "parametry":
                                int parametersCount = (int)((Excel.Range) worksheet.Cells[rowPos, column + 1]).Value2;
                                while (parametersCount-- > 0) {
                                    ++rowPos;
                                    parametersList.Add((string)((Excel.Range) worksheet.Cells[rowPos, column + 1]).Value2);
                                }
                                break;
                            case "wykres":
                                ++rowPos;
                                if (!reader.IsString(rowPos,column) || !reader.GetString(rowPos,column).Equals("nazwa:")) {
                                    error = true;
                                    errorString = "Błąd przy 'nazwa:' w 'wykres' - wiersz " + rowPos;
                                    break;
                                }

                                string name = reader.GetString(rowPos, column+1);

                                ++rowPos;
                                if (!reader.IsString(rowPos, column) || !reader.GetString(rowPos, column).Equals("dane:")) {
                                    error = true;
                                    errorString = "Błąd przy 'dane:' w 'wykres' - wiersz " + rowPos;
                                    break;
                                }

                                if (!reader.IsNumber(rowPos, column + 1)) {
                                    error = true;
                                    errorString = "Błąd przy 'dane:' w 'wykres' - wiersz " + rowPos;
                                    break;
                                }

                                int dataCount = reader.GetInt(rowPos, column + 1);

                                ++rowPos;
                                int seriesCount = 0;

                                while (reader.IsString(rowPos, column + seriesCount)) seriesCount++;

                                worksheet.Range[worksheet.Cells[rowPos, column], worksheet.Cells[rowPos + dataCount, column + seriesCount - 1]].Select();

                                rowPos += dataCount;
                                
                                DrawGraphFromSelectedCells(graphNum++, c, name, "", newWorksheet);


                                break;
                            case "dane":

                                Data data = new Data();

                                ++rowPos;
                                if (!reader.IsString(rowPos, column) || !reader.GetString(rowPos, column).Equals("parametry:") || !reader.IsNumber(rowPos, column + 1)) {
                                    error = true;
                                    errorString = "Błąd przy 'parametry:' w 'dane' - wiersz " + rowPos;
                                    break;
                                }

                                parametersCount = (int)((Excel.Range)worksheet.Cells[rowPos, column + 1]).Value2;

                                while (parametersCount-- > 0) {
                                    ++rowPos;
                                    data.Parameters.Add(reader.GetString(rowPos, column),reader.GetInt(rowPos,column+1));
                                }

                                ++rowPos;
                                if (!reader.IsString(rowPos, column) || !reader.GetString(rowPos, column).Equals("wartości:") || !reader.IsNumber(rowPos, column + 1)) {
                                    error = true;
                                    errorString = "Błąd przy 'wartości:' w 'dane' - wiersz " + rowPos;
                                    break;
                                }

                                int valuesCount = reader.GetInt(rowPos, column + 1);

                                while (valuesCount-- > 0) {
                                    ++rowPos;
                                    if (!reader.IsNumber(rowPos, column + 1)) {
                                        error = true;
                                        errorString = "Nie znaleziono liczby przy wartości - wiersz " + rowPos;
                                        break;
                                    }
                                    data.Values.Add(reader.GetString(rowPos, column), reader.GetDouble(rowPos, column + 1));
                                }

                                dataList.Add(data);

                                break;
                            case "tabela":

                                Table table = new Table();

                                ++rowPos;
                                if (!reader.IsString(rowPos, column) || !reader.GetString(rowPos, column).Equals("nazwa:")) {
                                    error = true;
                                    errorString = "Błąd przy 'nazwa:' w 'tabela' - wiersz " + rowPos;
                                    break;
                                }

                                name = reader.GetString(rowPos, column + 1);

                                table.Name = name;

                                ++rowPos;
                                if (!reader.IsString(rowPos, column) || !reader.GetString(rowPos, column).Equals("parametry:") || !reader.IsNumber(rowPos, column+1)) {
                                    error = true;
                                    errorString = "Błąd przy 'parametry:' w 'tabela' - wiersz " + rowPos;
                                    break;
                                }

                                parametersCount = (int)((Excel.Range)worksheet.Cells[rowPos, column + 1]).Value2;

                                while (parametersCount-- > 0) {
                                    ++rowPos;
                                    table.Parameters.Add(reader.GetString(rowPos, column));
                                }

                                ++rowPos;
                                if (!reader.IsString(rowPos, column) || !reader.GetString(rowPos, column).Equals("wartości:") || !reader.IsNumber(rowPos, column+1)) {
                                    error = true;
                                    errorString = "Błąd przy 'wartość:' w 'tabela' - wiersz "+rowPos;
                                    break;
                                }
                                
                                valuesCount = reader.GetInt(rowPos, column+1);

                                while (valuesCount-- > 0) {
                                    ++rowPos;
                                    table.Values.Add(reader.GetString(rowPos, column));
                                }

                                tables.Add(table);

                                break;
                        }

                        ++rowPos;

                        for (int rowsSearched = 0; rowsSearched < _MAX_SEARCH && ((Excel.Range)worksheet.Cells[rowPos, column]).Value2 == null
                            ; ++rowsSearched) ++rowPos;

                        if (((Excel.Range)worksheet.Cells[rowPos, column]).Value2 == null) foundNext = false;
                        
                    }

                    colTrans += 1;
                }

            }

            if (error) MessageBox.Show("błąd : " + errorString);

            int tableColumn = _GRAPH_WIDTH*columns+3;
            int tableRow = 1;

            foreach (Table table in tables) {

                //checking parameters range
                List<List<int>> parametersRanges = new List<List<int>>();
                foreach (string parameter in table.Parameters) {
                    List<int> range = new List<int>();
                    foreach (Data data in dataList) {
                        if (data.Parameters.ContainsKey(parameter)) {
                            int paramValue = data.Parameters[parameter];
                            if (!range.Contains(paramValue)) range.Add(paramValue);
                        }
                    }
                    range.Sort();
                    parametersRanges.Add(range);
                }

                if (table.Parameters.Count == 2) {

                    //title
                    ((Excel.Range) newWorksheet.Range[newWorksheet.Cells[tableRow, tableColumn], newWorksheet.Cells[tableRow, tableColumn + parametersRanges[0].Count + 1]]).Merge();
                    ((Excel.Range) newWorksheet.Cells[tableRow, tableColumn]).Value2 = table.Name;

                    //first parameter title
                    ((Excel.Range) newWorksheet.Range[newWorksheet.Cells[tableRow + 1, tableColumn + 2], newWorksheet.Cells[tableRow + 1, tableColumn + parametersRanges[0].Count + 1]]).Merge();
                    ((Excel.Range) newWorksheet.Cells[tableRow + 1, tableColumn + 2]).Value2 = table.Parameters[0];

                    //first parameter range
                    for (int i = 0; i < parametersRanges[0].Count; ++i) {
                        ((Excel.Range) newWorksheet.Cells[tableRow + 2, tableColumn + 2 + i]).Value2 = parametersRanges[0][i];
                    }

                    //second parameter title
                    ((Excel.Range) newWorksheet.Range[newWorksheet.Cells[tableRow + 3, tableColumn], newWorksheet.Cells[tableRow + 2 + parametersRanges[1].Count
                        , tableColumn]]).Merge();
                    ((Excel.Range) newWorksheet.Cells[tableRow + 3, tableColumn]).Value2 = table.Parameters[1];

                    //second parameter range
                    for (int i = 0; i < parametersRanges[1].Count; ++i) {
                        ((Excel.Range)newWorksheet.Cells[tableRow + 3 + i, tableColumn + 1]).Value2 = parametersRanges[1][i];
                    }

                    for (int v1 = 0;v1 < parametersRanges[0].Count;++v1) {
                        for (int v2 = 0; v2 < parametersRanges[1].Count; ++v2) {
                            bool found = false;
                            for (int d = 0; !found && d < dataList.Count; ++d) {
                                if (dataList[d].Parameters.ContainsKey(table.Parameters[0]) && dataList[d].Parameters[table.Parameters[0]] == parametersRanges[0][v1]
                                    && dataList[d].Parameters.ContainsKey(table.Parameters[1]) && dataList[d].Parameters[table.Parameters[1]] == parametersRanges[1][v2]
                                    && dataList[d].Values.ContainsKey(table.Values[0])) {
                                    found = true;
                                    ((Excel.Range)newWorksheet.Cells[tableRow + 3 + v2, tableColumn+2 + v1]).Value2 = dataList[d].Values[table.Values[0]];
                                }
                            }
                        }
                    }

                    tableRow += 3 + parametersRanges[1].Count + 1;

                } else if (table.Parameters.Count == 1) {

                    //title
                    ((Excel.Range)newWorksheet.Range[newWorksheet.Cells[tableRow, tableColumn], newWorksheet.Cells[tableRow, tableColumn + parametersRanges[0].Count]]).Merge();
                    ((Excel.Range)newWorksheet.Cells[tableRow, tableColumn]).Value2 = table.Name;

                    //Parameter title
                    ((Excel.Range)newWorksheet.Range[newWorksheet.Cells[tableRow + 1, tableColumn+1], newWorksheet.Cells[tableRow + 1, tableColumn + parametersRanges[0].Count]]).Merge();
                    ((Excel.Range)newWorksheet.Cells[tableRow + 1, tableColumn+1]).Value2 = table.Parameters[0];

                    //Parameter range
                    for (int i = 0; i < parametersRanges[0].Count; ++i) {
                        ((Excel.Range)newWorksheet.Cells[tableRow + 2, tableColumn + i + 1]).Value2 = parametersRanges[0][i];
                    }

                    //Values
                    for (int v = 0; v < table.Values.Count; ++v) {

                        //Value name
                        ((Excel.Range) newWorksheet.Cells[tableRow + 3 + v, tableColumn]).Value2 = table.Values[v];

                        for (int v1 = 0; v1 < parametersRanges[0].Count; ++v1) {
                            bool found = false;
                            for (int d = 0; !found && d < dataList.Count; ++d) {
                                if (dataList[d].Parameters.ContainsKey(table.Parameters[0]) && dataList[d].Parameters[table.Parameters[0]] == parametersRanges[0][v1]
                                    && dataList[d].Values.ContainsKey(table.Values[v])) {
                                    found = true;
                                    ((Excel.Range) newWorksheet.Cells[tableRow + 3 + v, tableColumn + v1 + 1]).Value2 = dataList[d].Values[table.Values[v]];
                                }
                            }
                        }

                    }

                    tableRow += 3 + table.Values.Count;

                }

            }

            

            Application.ScreenUpdating = screenUpdating;

        }

        public void DrawAllGraphs1Column(Office.CommandBarButton cmdBarbutton, ref bool cancel) {
            DrawAllGraphs(1);
        }

        public void DrawAllGraphs2Columns(Office.CommandBarButton cmdBarbutton, ref bool cancel) {
            DrawAllGraphs(2);
        }

        public Office.CommandBarButton AddOptionToContextMenu(String name, int beforePos
            , Office._CommandBarButtonEvents_ClickEventHandler handler) {

            Office.CommandBars commandbars = Globals.ThisAddIn.Application.CommandBars;

            Office.CommandBar ccb = commandbars["Cell"];
            Office.CommandBarButton button = (Office.CommandBarButton)ccb.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, beforePos + 1, true);
            button.Style = Office.MsoButtonStyle.msoButtonCaption;
            button.Caption = name;
            button.Tag = name;
            button.Visible = true;
            button.Click += handler;

            return button;

        }

        public Excel.Worksheet GetGraphsAndTablesWorksheet() {
            Excel.Worksheet newWorksheet = null;

            Worksheet activeWorksheet = (Worksheet)Globals.Factory.GetVstoObject(this.Application.ActiveWorkbook.ActiveSheet);

            foreach (Excel.Worksheet worksheet in Application.Worksheets) {
                if (worksheet.Name.Equals("Wykresy i Tabele")) {
                    newWorksheet = worksheet;
                }
            }


            if (newWorksheet == null) {
                newWorksheet = (Excel.Worksheet)Application.Worksheets.Add();
                newWorksheet.Move(After: Application.Worksheets.Item[Application.Worksheets.Count]);
                newWorksheet.Name = "Wykresy i Tabele";
                activeWorksheet.Activate();
            }

            return newWorksheet;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {

            //Dodajemy funkcje DrawGraph w menu kontekstowym
            _drawGraphButton = AddOptionToContextMenu("Rysuj Wykres", 0, DrawGraph);

            //Dodajemy funkcje DrawAllGraphs1Column
            _drawAllGraphsButton2 = AddOptionToContextMenu("Rysuj Wszystkie Wykresy (2 kolumny)", 0, DrawAllGraphs2Columns);

            //Dodajemy funkcje DrawAllGraphs1Column
            _drawAllGraphsButton1 = AddOptionToContextMenu("Rysuj Wszystkie Wykresy (1 kolumna)", 0, DrawAllGraphs1Column);

        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) {
            _drawGraphButton.Delete();
            _drawAllGraphsButton1.Delete();
            _drawAllGraphsButton2.Delete();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
