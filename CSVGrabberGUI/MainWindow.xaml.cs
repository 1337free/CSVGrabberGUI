using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Net;
using LumenWorks.Framework.IO.Csv;
using System.IO;
using Microsoft.Win32;

namespace WpfApplication1
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
			txtURL.Focus();
			try
			{
					RegistryKey SoftwareKey = Registry.LocalMachine.OpenSubKey("Software\\CSVGetter", false);
					txtURL.Text = (String) SoftwareKey.GetValue("URLLocation");
					txtOutput.Text = (String)SoftwareKey.GetValue("FileLocation");
					txtTime.Text = (String)SoftwareKey.GetValue("Time");
					SoftwareKey.Close();
					MessageBox.Show("Previously saved data has automatically been loaded.", "Previous data loaded", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception)
			{

			}

		}

		private void btnTest_Click(object sender, RoutedEventArgs e)
		{
				if (!Validate())
				{ return; }

				WebClient Client = new WebClient();
				try
				{
				Client.DownloadFile(txtURL.Text, System.IO.Path.GetTempPath().ToString() + "temp.csv");
				}
				catch (Exception)
				{
					MessageBox.Show("Error! The entered URL: \"" + txtURL.Text + "\" is not valid or you do not have sufficient permissions to access it.", "Invalid Url", MessageBoxButton.OK, MessageBoxImage.Error);
					return;
				}
				Client.DownloadFile(txtURL.Text, System.IO.Path.GetTempPath().ToString() + "temp.csv");
				Client.Dispose();

				StreamReader strmReader;
				StreamWriter strmWriter;
				CsvReader strmCSVReader;

				try
				{
					strmReader = File.OpenText(System.IO.Path.GetTempPath().ToString() + "temp.csv");
					strmWriter = new StreamWriter(txtOutput.Text);
					strmCSVReader = new CsvReader(strmReader, true);

					foreach (String[] value in strmCSVReader.ToArray())
					{
							strmWriter.WriteLine(value[0]);
					}

					strmReader.Close();
					strmWriter.Close();
					strmCSVReader.Dispose();

				}
						catch (Exception)
			{
        MessageBox.Show("Error! The output location: \"" + txtOutput.Text + "\" is not valid or you do not have sufficient permissions to access it.", "Invalid Output Location", MessageBoxButton.OK, MessageBoxImage.Error);
        return;	
				}

				MessageBox.Show("The input CSV was succesfully parsed and saved to: " + txtOutput.Text, "Conversion Successful!");
		}

		private void btnSave_Click(object sender, RoutedEventArgs e)
		{

				if (!Validate())
				{ return; }

				RegistryKey SoftwareKey = Registry.LocalMachine.OpenSubKey("Software", true);
				SoftwareKey = SoftwareKey.CreateSubKey("CSVGetter", RegistryKeyPermissionCheck.ReadWriteSubTree);
				SoftwareKey.SetValue("URLLocation", txtURL.Text, RegistryValueKind.String);
				SoftwareKey.SetValue("FileLocation", txtOutput.Text, RegistryValueKind.String);
				SoftwareKey.SetValue("Time", txtTime.Text, RegistryValueKind.String);
				SoftwareKey.Close();
        System.ServiceProcess.ServiceController sc = new System.ServiceProcess.ServiceController("CSVGrabber");

        try
        {
          if (sc.Status == System.ServiceProcess.ServiceControllerStatus.Running) {
          sc.Stop();
          sc.WaitForStatus(System.ServiceProcess.ServiceControllerStatus.Stopped);
          }
          sc.Start();      
          MessageBox.Show("Data has been saved successfully to the registry. And the CSVGrabber service has been restarted.", "Data Saved Successfully", MessageBoxButton.OK, MessageBoxImage.Information);

        }
        catch (Exception)
        {
          MessageBox.Show("Data has been saved successfully to the registry.", "Data Saved Successfully", MessageBoxButton.OK, MessageBoxImage.Information);
          
        }
        finally {
          sc.Dispose();          
        }
								
		}

		private Boolean Validate ()
		{

			if (txtURL.Text.Length == 0 || txtOutput.Text.Length == 0 || txtTime.Text.Length == 0)
			{
				MessageBox.Show("No empty fields are allowed.", "Empty Field", MessageBoxButton.OK, MessageBoxImage.Error);
				return false; 
			}

			try 
	{
		DateTime.ParseExact(txtTime.Text, "HHmm", System.Globalization.CultureInfo.InvariantCulture);
	}
	catch (Exception)
	{
		MessageBox.Show("Please enter a valid 24 hour time to run in the HHmm format.", "Invalid Time", MessageBoxButton.OK, MessageBoxImage.Error);
		return false; 
	}
			
				return true;
		}

		private void btnBrowse_Click(object sender, RoutedEventArgs e)
		{

			Microsoft.Win32.SaveFileDialog dlgBrowse = new Microsoft.Win32.SaveFileDialog();
			dlgBrowse.DefaultExt = ".txt";
			dlgBrowse.Filter = "Text Documents|*.txt|All Files|*.*";
			if (!(bool)dlgBrowse.ShowDialog())
			{
				return;
			}
			txtOutput.Text = dlgBrowse.FileName;

		}

		private void btnClearData_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				Registry.LocalMachine.DeleteSubKey("Software\\CSVGetter");
				txtURL.Clear();
				txtOutput.Clear();
				txtTime.Clear();
				MessageBox.Show("Previously saved data has been cleared.", "Previous data cleared", MessageBoxButton.OK, MessageBoxImage.Exclamation);
			}
			catch (Exception)
			{
				MessageBox.Show("There is no previously saved data to clear, or you do not have sufficient privileges.", "Previous data not cleared", MessageBoxButton.OK, MessageBoxImage.Error);
			}

		}

    private void clickAbout(object sender, MouseEventArgs e)
    {
      MessageBox.Show("CSVGrabber GUI 0.002 Written by Andrew Zumhagen © " + DateTime.Now.ToString() + "(AKA, right now...)\n\nFor Evan Sink. May your sink never sink.", "About", MessageBoxButton.OK, MessageBoxImage.Information);
    }

	}
}

