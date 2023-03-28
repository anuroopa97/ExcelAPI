using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelPI
{
    public class Session
    {
        /// <summary>
        /// Session API
        /// </summary>
        public List<Workbook> Workbooks { get; private set; }

        /// <summary>
        /// Initializes a new instance of workbook.
        /// </summary>
        public Session()
        {
            Workbooks = new List<Workbook>();
        }

        /// <summary>
        /// Opens a workbook from given path
        /// </summary>
        public Workbook OpenWorkbook(string path)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine($"Error: File '{path}' not found.");
                return null;
            }

            var workbook = new Workbook(this, path);
            Workbooks.Add(workbook);
            workbook.Open();

            return workbook;
        }

        /// <summary>
        /// Creates a new, empty workbook
        /// </summary>
        public Workbook CreateWorkbook(string name)
        {
            var workbook = new Workbook(this, name);
            Workbooks.Add(workbook);

            return workbook;
        }

        /// <summary>
        /// Closes the specified workbook.
        /// </summary>
        public void CloseWorkbook(Workbook workbook)
        {
            Workbooks.Remove(workbook);
            workbook.Close();
        }
    }

    /// <summary>
    /// Excel WorkBook API.
    /// </summary>
    public class Workbook
    {
        private Session _session;

        /// <summary>
        /// Gets the list of worksheets in this workbook.
        /// </summary>
        public List<Worksheet> Worksheets { get; private set; }

        /// <summary>
        /// Gets or sets the name of this workbook.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the file path of this workbook.
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Initializes a new instance
        /// </summary>
        public Workbook(Session session, string path)
        {
            _session = session;
            Path = path;
            Name = System.IO.Path.GetFileName(path);
            Worksheets = new List<Worksheet>();
        }
        /// <summary>
        /// Opens this workbook.
        /// </summary>
        public void Open()
        {
            Console.WriteLine($"Opening workbook '{Name}' from file '{Path}'...");
            // implementation to open the workbook from file
        }
        /// <summary>
        /// Saves this workbook.
        /// </summary>
        public void Save()
        {
            Console.WriteLine($"Saving workbook '{Name}' to file '{Path}'...");
            // implementation to save the workbook to file
        }

        /// <summary>
        /// Saves this workbook
        /// </summary>
        public void SaveAs(string newPath)
        {
            Console.WriteLine($"Saving workbook '{Name}' to new file '{newPath}'...");
            // implementation to save the workbook to the new file path
            Path = newPath;
        }

        /// <summary>
        /// Closes this workbook.
        /// </summary>
        public void Close()
        {
            Console.WriteLine($"Closing workbook '{Name}'...");
            // implementation to close the workbook
        }

        /// <summary>
        /// Adds a new worksheet
        /// </summary>
        public Worksheet AddWorksheet(string name)
        {
            var worksheet = new Worksheet(this, name);
            Worksheets.Add(worksheet);

            return worksheet;
        }

        /// <summary>
        /// to delete worksheet
        /// </summary>
        public void DeleteWorksheet(Worksheet worksheet)
        {
            Worksheets.Remove(worksheet);
        }

        /// <summary>
        /// rename functionality
        /// </summary>
        public void RenameWorksheet(Worksheet worksheet, string newName)
        {
            worksheet.Name = newName;
        }

        /// <summary>
        /// activating a worksheet
        /// </summary>
        public void ActivateWorksheet(Worksheet worksheet)
        {
            // implementation to activate the specified worksheet
        }
    }

    /// <summary>
    /// Worksheet API
    /// </summary>
    public class Worksheet
    {
        private Workbook _workbook;

        /// <summary>
        /// getter and setter for the worksheet
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// new instance can be initialized here <see cref="Worksheet"/>
        /// </summary>
        public Worksheet(Workbook workbook, string name)
        {
            _workbook = workbook;
            Name = name;
        }
    }
}
