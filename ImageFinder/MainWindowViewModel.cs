using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Input;
using JetBrains.Annotations;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace ImageFinder
{
    public class ObservableObject : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }

    public class DelegateCommand : ICommand
    {
        private readonly Action _action;

        public DelegateCommand(Action action)
        {
            _action = action;
        }

        public void Execute(object parameter)
        {
            _action();
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public event EventHandler CanExecuteChanged;
    }

    public class DelegateCommand<T> : ICommand
    {
        private readonly Action<T> _action;

        public DelegateCommand(Action<T> action)
        {
            _action = action;
        }

        public void Execute(object parameter)
        {
            _action((T) parameter);
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public event EventHandler CanExecuteChanged;
    }

    public class MainWindowViewModel : ObservableObject
    {
        private const string AccessString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Mode=Read";
        private const string ExcelString = @"Provider=Microsoft.ACE.OLEDB.12.0;Mode=Read;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""";
        private Dictionary<string, string> _fileCache;
        private string _nameFile;
        public string DatabaseFile { get; private set; }

        public ICommand OpenDatabaseFileCommand
        {
            get { return new DelegateCommand(OpenDatabaseFile); }
        }

        public ICommand OpenDirectoryCommand
        {
            get { return new DelegateCommand(OpenDirectory); }
        }

        public DataView Names { get; set; }

        public string NameFile
        {
            get { return _nameFile; }
            private set
            {
                var connectionString = string.Format(ExcelString, value);
                OleDbConnection connection;
                try
                {
                    connection = new OleDbConnection(connectionString);
                }
                catch (ArgumentException)
                {
                    return;
                }
                try
                {
                    connection.Open();
                    var sheet = GetFirstTable(connection);
                    var selectCommandText = string.Format("SELECT * FROM [{0}]", sheet);
                    var adapter = new OleDbDataAdapter(selectCommandText, connection);
                    var ds = new DataSet();
                    adapter.Fill(ds, "Names");
                    Names = ds.Tables["Names"].AsDataView();
                }
                catch (OleDbException)
                {
                    return;
                }
                finally
                {
                    connection.Dispose();
                }
                _nameFile = value;
                OnPropertyChanged();
                OnPropertyChanged("Names");
            }
        }

        public ICommand OpenImageCommand
        {
            get { return new DelegateCommand<DataRowView>(OpenImage); }
        }

        public ICommand OpenNameFileCommand
        {
            get { return new DelegateCommand(OpenNameFile); }
        }

        public string DirectoryPath { get; private set; }

        private static string GetFirstTable(OleDbConnection connection)
        {
            var tables = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {null, null, null, "TABLE"});
            var sheet = tables.Rows[0].Field<string>("TABLE_NAME");
            return sheet;
        }

        private static IEnumerable<string> DirSearch(string dir)
        {
            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(dir);
            }
            catch (UnauthorizedAccessException)
            {
                yield break;
            }
            foreach (var file in files)
            {
                yield return file;
            }
            foreach (var file in Directory.EnumerateDirectories(dir).SelectMany(DirSearch))
            {
                yield return file;
            }
        }

        private static string GetImageName(OleDbConnection connection, string table, string column, string name)
        {
            for (var i = 2; i > 0; i--)
            {
                var query = string.Format("SELECT IMAGE_ID FROM {0} WHERE {1} LIKE '{2}%' AND MULTI_RECORD_TYPE = 'Type {3}'", table, column, name, i);
                using (var command = new OleDbCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        if (reader == null || !reader.Read())
                        {
                            continue;
                        }
                        var imagePath = reader[0].ToString();
                        var imageName = Path.GetFileName(imagePath);
                        return imageName;
                    }
                }
            }
            return null;
        }

        private void OpenImage(DataRowView row)
        {
            if (DatabaseFile == null || DirectoryPath == null)
            {
                return;
            }
            var name = row.Row[0].ToString();
            var connectionString = string.Format(AccessString, DatabaseFile);
            OleDbConnection dbConnection;
            using (dbConnection = new OleDbConnection(connectionString))
            {
                dbConnection.Open();
                var table = GetFirstTable(dbConnection);
                var columnName = Names.Table.Columns[0].ColumnName;
                var isColumnPresent = dbConnection.GetSchema("Columns").AsEnumerable().Any(a => a.Field<string>("TABLE_NAME") == table && a.Field<string>("COLUMN_NAME") == columnName);
                if (!isColumnPresent)
                {
                    MessageBox.Show(string.Format("Can't find '{0}' in database", columnName));
                    return;
                }
                var imageName = GetImageName(dbConnection, table, columnName, name);
                if (imageName == null)
                {
                    MessageBox.Show(string.Format("Can't find '{0}' in database", name));
                    return;
                }
                string image;
                if (_fileCache.TryGetValue(imageName, out image))
                {
                    Process.Start(image);
                }
                else
                {
                    MessageBox.Show(string.Format("Can't find '{0}' in directory", imageName));
                }
            }
        }

        private void OpenDatabaseFile()
        {
            var dialog = new OpenFileDialog {Filter = "Access File|*.accdb"};
            if (dialog.ShowDialog() == true)
            {
                var dbFile = dialog.FileName;
                var connectionString = string.Format(AccessString, dbFile);
                OleDbConnection connection;
                try
                {
                    connection = new OleDbConnection(connectionString);
                }
                catch (ArgumentException)
                {
                    return;
                }
                try
                {
                    connection.Open();
                }
                catch (OleDbException)
                {
                    return;
                }
                finally
                {
                    connection.Dispose();
                }
                DatabaseFile = dialog.FileName;
                OnPropertyChanged("DatabaseFile");
            }
        }

        private void OpenDirectory()
        {
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                DirectoryPath = dialog.SelectedPath;
                var fc = new Dictionary<string, string>();
                foreach (var file in DirSearch(DirectoryPath))
                {
                    var name = Path.GetFileName(file);
                    if (name != null)
                    {
                        fc[name] = file;
                    }
                }
                _fileCache = fc;
                OnPropertyChanged("DirectoryPath");
            }
        }

        private void OpenNameFile()
        {
            var dialog = new OpenFileDialog {Filter = "Excel File|*.xlsx"};
            if (dialog.ShowDialog() == true)
            {
                NameFile = dialog.FileName;
            }
        }
    }
}