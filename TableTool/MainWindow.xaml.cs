using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;

namespace TableTool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private class ExcelTableFilePath
        {
            public ExcelTableFilePath(string filePath)
            {
                this.filePath = filePath;
            }
            public string filePath { get; set; }

        }
        private List<ExcelTableFilePath> mFilePathList = new List<ExcelTableFilePath>();
        private List<string> mSelectedList = new List<string>();

        public List<T> GetChildObjects<T>(DependencyObject obj, string name) where T : FrameworkElement
        {
            DependencyObject child = null;
            List<T> childList = new List<T>();
            for (int i = 0; i <= VisualTreeHelper.GetChildrenCount(obj) - 1; i++)
            {
                child = VisualTreeHelper.GetChild(obj, i);
                if (child is T && (((T)child).Name == name || string.IsNullOrEmpty(name)))
                {
                    childList.Add((T)child);
                }
                childList.AddRange(GetChildObjects<T>(child, name));//指定集合的元素添加到List队尾
            }
            return childList;
        }

        private void AddSelected(string filePath)
        {
            if (mSelectedList.Contains(filePath))
                return;
            mSelectedList.Add(filePath);
        }

        private void RemoveSelected(string filePath)
        {
            if (!mSelectedList.Contains(filePath))
                return;
            mSelectedList.Remove(filePath);
        }

        private void RefreshCheckBox()
        {
            List<CheckBox> collection = GetChildObjects<CheckBox>(this.listView1, "SelectCheckBox");
            foreach (var v in collection)
            {
                bool ckecked = mSelectedList.Contains(v.Tag.ToString());
                v.IsChecked = ckecked;
            }
            bool allChecked = mSelectedList.Count != 0 && mSelectedList.Count == mFilePathList.Count;
            this.AllSelectCheckBox.IsChecked = allChecked;
        }

        public MainWindow()
        {
            InitializeComponent();
            this.listView1.ItemsSource = mFilePathList;
        }

        private void Button_Click_Generate(object sender, RoutedEventArgs e)
        {
            if (mFilePathList.Count == 0)
            {
                MessageBox.Show("没有xls文件。", "");
                return;
            }
            string[] filePaths = new string[mFilePathList.Count];
            for (int i = 0; i < filePaths.Length; ++i)
            {
                filePaths[i] = mFilePathList[i].filePath;
            }
            CodeType codeTypes = CodeType.NULL;
            if (U3dCsCheckBox.IsChecked.HasValue && U3dCsCheckBox.IsChecked.Value)
                codeTypes |= CodeType.U3D_CS;
            if (JavaCheckBox.IsChecked.HasValue && JavaCheckBox.IsChecked.Value)
                codeTypes |= CodeType.JAVA;
            if (CppCheckBox.IsChecked.HasValue && CppCheckBox.IsChecked.Value)
                codeTypes |= CodeType.CPP;
            OutputTable.OneKeyBuild(filePaths, codeTypes);
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            var data = (System.Array)e.Data.GetData(DataFormats.FileDrop);
            foreach (var v in data)
            {
                string fileName = v.ToString();
                if (System.IO.Path.GetExtension(fileName).ToLower() != ".xls")
                    continue;
                if (!mFilePathList.Exists(x => x.filePath == fileName))
                {
                    mFilePathList.Add(new ExcelTableFilePath(fileName));
                    this.listView1.Items.Refresh();
                }
                Console.WriteLine(fileName);
            }
            if (mFilePathList.Count > 0)
                PromptLabel.Visibility = System.Windows.Visibility.Hidden;
        }

        private void AllSelectCheckBox_Click(object sender, RoutedEventArgs e)
        {
            CheckBox allSelect = sender as CheckBox;
            mSelectedList.Clear();
            if (allSelect.IsChecked.Value)
            {
                foreach (var v in mFilePathList)
                    mSelectedList.Add(v.filePath);
            }
            RefreshCheckBox();
        }

        private void SelectCheckBox_Click(object sender, RoutedEventArgs e)
        {
            CheckBox check = sender as CheckBox;
            if (check.IsChecked.Value)
            {
                AddSelected(check.Tag.ToString());
            }
            else
            {
                RemoveSelected(check.Tag.ToString());
            }
            RefreshCheckBox();
        }

        private void Button_Click_Delete(object sender, RoutedEventArgs e)
        {
            if (mSelectedList.Count == 0)
                return;
            foreach (var v in mSelectedList)
            {
                mFilePathList.RemoveAll(x => x.filePath == v);
            }
            mSelectedList.Clear();
            this.listView1.Items.Refresh();
            RefreshCheckBox();
            if (mFilePathList.Count == 0)
                PromptLabel.Visibility = System.Windows.Visibility.Visible;
        }
    }
}
