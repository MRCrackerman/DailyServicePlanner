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

namespace _10Tim;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    ExcelHandler excel;
    ExecutionHandler exe;

    public MainWindow()
    {
        Init();
        //exe.CreateMonth();
        InitializeComponent();

        
    }

    private void Init(){
        excel = new ExcelHandler();
        exe = new ExecutionHandler();
    }

    private void RunMethodButton_Click(object sender, RoutedEventArgs e)
    {   
        //exe.CreateMonth();
        //MessageBox.Show("10 τιμητικες στον καθένα");
        try{
            //MessageBox.Show("prospatho na trekso");

            exe.CreateMonth();
        }
        catch(Exception ex){
            MessageBox.Show(ex.ToString());
        }
    }
}