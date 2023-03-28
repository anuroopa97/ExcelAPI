using System;
using System.Collections.Generic;

namespace ExcelAPI
{
    public class Session
    {
        private List<Workbook> _workbooks;

        public Session()
        {
            _workbooks = new List<Workbook>();
        }

        public Workbook OpenWorkbook(string path)
        {
            var workbook = new Workbook(path);
            _workbooks.Add(workbook);
            return workbook;
        }

        public void CloseWorkbook(Workbook workbook)
        {
            _workbooks.Remove(workbook);
            workbook.Close();
        }
    }

    public class Workbook
    {
        private string _path;
        private List<Worksheet> _worksheets;

        public Workbook(string path)
        {
            _path = path;
            _worksheets = new List<Worksheet>();
        }

        public Worksheet AddWorksheet(string name)
        {
            if (_worksheets.Exists(ws => ws.Name == name))
            {
                throw new ArgumentException("Worksheet with the same name already exists.");
            }

            var worksheet = new Worksheet(name);
            _worksheets.Add(worksheet);
            return worksheet;
        }
        public void Open()
        {
            // Close the workbook and release any resources.
        } 
        public void Save()
        {
            // Close the workbook and release any resources.
        } 
        public void SaveAs(string name)
        {
            // Close the workbook and release any resources.
        }
        public void Close()
        {
            // Close the workbook and release any resources.
        }

    }

    public class Worksheet
    {
        public string Name { get; }

        public Worksheet(string name)
        {
            Name = name;
        }

        public void MakeActive()
        {
            // Make this worksheet the active one.
        }

        public void Delete()
        {
            // Delete this worksheet.
        }

        public void Rename(string name)
        {
            // Rename this worksheet to the specified name.
        }
    }
}