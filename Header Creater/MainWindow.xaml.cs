using System.Linq;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System;


namespace Header_Creater
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
            date.Text = DateTime.Now.ToShortDateString();
            time.Text = DateTime.Now.ToString("hh: mm :ss tt");
            pBar1.Value = 0;
            Button1.IsEnabled = false;
            Button2.IsEnabled = false;
            Button3.IsEnabled = false;
            Button4.IsEnabled = false;
            Button5.IsEnabled = false;
            Button6.IsEnabled = false;
            Button7.IsEnabled = false;
            Button8.IsEnabled = false;
            Createfolder.IsEnabled = false;
            feedback.IsEnabled = false;
            olddyno.Items.Clear();
            olddyno.Items.Add("0258");
            olddyno.Items.Add("5419");
            olddyno.Items.Add("3362");
            olddyno.Items.Add("3084");
            olddyno.Items.Add("3088");
            olddyno.Items.Add("3107");
            olddyno.Items.Add("5577");
            olddyno.Items.Add("3165");
            olddyno.Items.Add("3186");
            newdyno.Items.Clear();
            newdyno.Items.Add("0258");
            newdyno.Items.Add("5419");
            newdyno.Items.Add("3362");
            newdyno.Items.Add("3084");
            newdyno.Items.Add("3088");
            newdyno.Items.Add("3107");
            newdyno.Items.Add("5577");
            newdyno.Items.Add("3165");
            newdyno.Items.Add("3186");
            reassign.IsEnabled = false;
            Testlistnew.IsEnabled = false;
            if (Directory.Exists("K:\\"))
            {
                Selecttest.Items.Clear();
                var files = Directory.EnumerateFiles(@"K:\\NewTestRequests\\")
                                .Where(file => file.ToLower().EndsWith("xlsx")
                                       || file.ToLower().EndsWith("xlsm"));
                foreach (var file in files)
                {
                    Selecttest.Items.Add(System.IO.Path.GetFileNameWithoutExtension(file));
                    feedback.Text = "Select Test";
                }
            }
            else
            {
                feedback.Text = "Please connect to K Drive.";
            }
        }
        private void Selecttest_TextUpdate(object sender, EventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string text = File.ReadAllText("K:\\Header Default Setup\\DSH\\" + Dyno.Text + ".txt");
            MakeHeader.IsEnabled = true;
            Button9.IsEnabled = true;
            text = text.Replace("test ", "test " + Testnum.Text)
            .Replace("407 -1 ", "407 -1 " + Testnum.Text)
            .Replace("400 -1 ", "400 -1 " + date.Text)
            .Replace("401 -1 ", "401 -1 " + time.Text)
            .Replace("408 -1 ", "408 -1 " + TScript.Text)
            .Replace("409 -1 ", "409 -1 " + TSpec.Text)
            .Replace("410 -1 ", "410 -1 " + projectnumber.Text)
            .Replace("415 -1 ", "415 -1 " + Submittedby.Text)
            .Replace("413 -1 ", "413 -1 " + owner.Text)
            .Replace("422 -1 ", "422 -1 " + reason.Text)
            .Replace("428 -1 ", "428 -1 " + projectname.Text)
            .Replace("442 -1 ", "442 -1 " + Brakename.Text)
            .Replace("500 51 ", "500 51 " + rollingrad.Text)
            .Replace("501 61 ", "501 61 " + reqinertia.Text)
            .Replace("510 51 ", "510 51 " + pistondia.Text)
            .Replace("511 51 ", "511 51 " + effradius.Text)
            .Replace("530 2 ", "530 2 " + pistonnum.Text)
            .Replace("674 -1 ", "674 -1 " + knuckle.Text)
            .Replace("696 -1 ", "696 -1 " + driveadapt.Text)
            .Replace("433 -1 ", "433 -1 " + rotorid.Text)
            .Replace("432 -1 ", "432 -1 " + caliper.Text)
            .Replace("434 -1 ", "434 -1 " + rotorsize.Text)
            .Replace("454 -1 ", "454 -1 " + rotorsource.Text)
            .Replace("416 -1 ", "416 -1 " + batchinner.Text)
            .Replace("414 -1 ", "414 -1 " + batchouter.Text)
            .Replace("417 -1 ", "417 -1 " + Materialsalesinner.Text)
            .Replace("418 -1 ", "418 -1 " + Materialsalesouter.Text)
            .Replace("420 -1 ", "420 -1 " + TestLog.Text)
            .Replace("436 -1 1", "436 -1 " + padinner.Text)
            .Replace("438 -1 3", "438 -1 " + padouter.Text)
            .Replace("681 -1 ", "681 -1 " + underlayinner.Text)
            .Replace("682 -1 ", "682 -1 " + underlayouter.Text)
            .Replace("679 -1 ", "679 -1 " + insulatorinner.Text)
            .Replace("680 -1 ", "680 -1 " + insulatorouter.Text)
            .Replace("683 -1 ", "683 -1 " + chamferinner.Text)
            .Replace("684 -1 ", "684 -1 " + chamferouter.Text)
            .Replace("685 -1 ", "685 -1 " + slotinner.Text)
            .Replace("686 -1 ", "686 -1 " + slotouter.Text)
            .Replace("691 -1 ", "689 -1 " + compinner.Text)
            .Replace("692 -1 ", "692 -1 " + compouter.Text)
            .Replace("695 -1 ", "695 -1 " + anchor.Text)
            .Replace("748 -1 ", "748 -1 " + testeng.Text);
            if (Directory.Exists(@"K:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno operations"))
            {
                File.WriteAllText("K:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno operations\\" + Selecttest.SelectedValue + ".DSH", text);
                Header.Text = File.ReadAllText("K:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno operations\\" + Selecttest.SelectedValue + ".DSH");
                File.WriteAllText("K:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno operations\\" + Selecttest.SelectedValue + ".DSH", Header.Text);
                MakeHeader.IsEnabled = false;
                Button9.IsEnabled = true;
                feedback.Text = "Header Created Successfully. for further changes manually change header below then press update Header";
            }
            else
            {
                feedback.Text = "A Folder isn't created for this test please Create Test folder.";
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Directory.CreateDirectory(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\3D scan");
            Directory.CreateDirectory(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\DTV");
            Directory.CreateDirectory(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno Operations");
            Directory.CreateDirectory(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\FoX TR");
            Directory.CreateDirectory(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Grindo");
            Directory.CreateDirectory(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Pictures");
            Directory.CreateDirectory(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Report");
            Directory.CreateDirectory(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\surfs");
            File.Copy(@"K:\Header Default Setup\HCF\FM 2019 Header Config - " + Dyno.Text + ".HCF", @"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno Operations\\FM 2019 Header Config - " + Dyno.Text + ".HCF");
            File.Move(@"K:\\NewTestRequests\\" + Selecttest.SelectedValue + ".xlsx", @"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\FoX TR\\" + Selecttest.SelectedValue + ".xlsx");
            feedback.Text = "Test Folder Created For Test: " + Selecttest.SelectedItem + " In Dyno: " + Dyno.Text + ", Checkout script.";
            Button8.IsEnabled = true;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            TabControl1.SelectedIndex = 0;
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            TabControl1.SelectedIndex = 1;
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            TabControl1.SelectedIndex = 2;
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            TabControl1.SelectedIndex = 3;
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            TabControl1.SelectedIndex = 4;
        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            TabControl1.SelectedIndex = 5;
        }

        private void Button7_Click(object sender, RoutedEventArgs e)
        {
            TabControl1.SelectedIndex = 6;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(Testnum.Text))
            {
                ListBox2.Items.Add("Test Number is empty");
            }
            else
            {
                ListBox2.Items.Remove("Test Number is empty");
            }
        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(pistonnum.Text))
            {
                ListBox2.Items.Add("Piston Number is empty");
            }
            else
            {
                ListBox2.Items.Remove("Piston Number is empty");
            }
        }

        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            if (File.Exists(@"K:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno operations\\" + Selecttest.SelectedValue + ".DSH"))
            {
                File.WriteAllText("K:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno operations\\" + Selecttest.SelectedValue + ".DSH", Header.Text);
                feedback.Text = "Header updated";
            }
            else
            {
                File.WriteAllText("K:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue + "\\Dyno operations\\" + Selecttest.SelectedValue + ".DSH", Header.Text);
                feedback.Text = "Header updated";
            }
        }
        private void Button_Click_9(object sender, RoutedEventArgs e)
        {

            Selecttest.Items.Clear();
            if (Directory.Exists("k:\\"))
            {


                var files = Directory.EnumerateFiles(@"K:/NewTestRequests/")
                            .Where(file => file.ToLower().EndsWith("xlsx")
                                   || file.ToLower().EndsWith("xlsm"));
                foreach (var file in files)
                {
                    Selecttest.Items.Add(System.IO.Path.GetFileNameWithoutExtension(file));
                    for (int n = Selecttest.Items.Count - 1; n >= 0; --n)
                    {
                        string removelistitem = "~$";
                        if (Selecttest.Items[n].ToString().Contains(removelistitem))
                        {
                            Selecttest.Items.RemoveAt(n);
                        }
                    }
                    feedback.Text = "List Refreshed, Please select Test.";
                }
            }
            else
            {
                feedback.Text = "This Device isn't connected to K Drive.";

            }
        }

        private void TabControl1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Rotorid_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(rotorid.Text))
            {
                ListBox2.Items.Add("Rotor ID is empty");
            }
            else
            {
                ListBox2.Items.Remove("Rotor ID is empty");
            }

        }


        private void Button_Click_10(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(@"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue))
            {
                string path = @"k:\\Dyno Testing\\" + Dyno.Text + "\\" + Selecttest.SelectedValue;
                System.Diagnostics.Process.Start(path);
            }
            else
            {
                feedback.Text = "Test Folder Doesn't exist, Create Test Folder first";
            }
        }

        private void Dyno_TextChanged(object sender, TextChangedEventArgs e)
        {

            if (String.IsNullOrEmpty(Dyno.Text))
            {
                ListBox2.Items.Add("Important!, Dyno is empty");
                Createfolder.IsEnabled = false;
                Button8.IsEnabled = false;
                feedback.Text = "Dyno is empty Cannot use program if no Dyno is selected";
            }
            else
            {
                ListBox2.Items.Remove("Important!, Dyno is empty");
                Createfolder.IsEnabled = true;
                Button8.IsEnabled = true;
            }


        }

        private void TestLog_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(TestLog.Text))
            {
                ListBox2.Items.Add("Test Log is empty");
            }
            else
            {
                ListBox2.Items.Remove("Test Log is empty");
            }
        }

        private void TScript_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(TScript.Text))
            {
                ListBox2.Items.Add("Test Script is empty");
            }
            else
            {
                ListBox2.Items.Remove("Test Script is empty");
            }
        }

        private void TSpec_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(TSpec.Text))
            {
                ListBox2.Items.Add("Test Specification is empty");
            }
            else
            {
                ListBox2.Items.Remove("Test Specification is empty");
            }
        }

        private void Submittedby_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(Submittedby.Text))
            {
                ListBox2.Items.Add("Submitted by is empty");
            }
            else
            {
                ListBox2.Items.Remove("Submitted by is empty");
            }
        }

        private void Owner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(owner.Text))
            {
                ListBox2.Items.Add("Owner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Owner is empty");
            }
        }

        private void Projectname_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(projectname.Text))
            {
                ListBox2.Items.Add("Project Name is empty");
            }
            else
            {
                ListBox2.Items.Remove("Project Name is empty");
            }
        }

        private void Projectnumber_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(projectnumber.Text))
            {
                ListBox2.Items.Add("Project Number is empty");
            }
            else
            {
                ListBox2.Items.Remove("Project Number is empty");
            }
        }

        private void Reason_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(reason.Text))
            {
                ListBox2.Items.Add("Reason is empty");
            }
            else
            {
                ListBox2.Items.Remove("Reason is empty");
            }
        }

        private void Pistondia_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(pistondia.Text))
            {
                ListBox2.Items.Add("Piston Diameter is empty");
            }
            else
            {
                ListBox2.Items.Remove("Piston Diameter is empty");
            }
        }

        private void Rollingrad_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(rollingrad.Text))
            {
                ListBox2.Items.Add("Rolling Radius is empty");
            }
            else
            {
                ListBox2.Items.Remove("Rolling Radius is empty");
            }
        }

        private void Effradius_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(effradius.Text))
            {
                ListBox2.Items.Add("Effective Radius is empty");
            }
            else
            {
                ListBox2.Items.Remove("Effective Radius is empty");
            }
        }

        private void Reqinertia_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(reqinertia.Text))
            {
                ListBox2.Items.Add("Required Inertia is empty");
            }
            else
            {
                ListBox2.Items.Remove("Required Inertia is empty");
            }
        }

        private void Brakename_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(Brakename.Text))
            {
                ListBox2.Items.Add("Brake Name is empty");
            }
            else
            {
                ListBox2.Items.Remove("Brake Name is empty");
            }
        }

        private void Rotorsize_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(rotorsize.Text))
            {
                ListBox2.Items.Add("Rotor Size is empty");
            }
            else
            {
                ListBox2.Items.Remove("Rotor Size is empty");
            }
        }

        private void Rotorsource_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(rotorsource.Text))
            {
                ListBox2.Items.Add("Rotor Source is empty");
            }
            else
            {
                ListBox2.Items.Remove("Rotor Source is empty");
            }
        }

        private void Fixture_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void FixtureType_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Driveadapt_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Knuckle_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Caliper_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Anchor_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Batchinner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(batchinner.Text))
            {
                ListBox2.Items.Add("Batch Inner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Batch inner is empty");
            }
        }

        private void Batchouter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(batchouter.Text))
            {
                ListBox2.Items.Add("Batch outer is empty");
            }
            else
            {
                ListBox2.Items.Remove("Batch Outer is empty");
            }
        }

        private void Materialsalesinner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(Materialsalesinner.Text))
            {
                ListBox2.Items.Add("Material Sales Inner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Material Sales Inner is empty");
            }
        }

        private void Materialsalesouter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(Materialsalesouter.Text))
            {
                ListBox2.Items.Add("Material Sales Outer is empty");
            }
            else
            {
                ListBox2.Items.Remove("Material Sales Outer is empty");
            }

        }

        private void Padinner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(padinner.Text))
            {
                ListBox2.Items.Add("Pad Inner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Pad Inner is empty");
            }

        }

        private void Padouter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(padouter.Text))
            {
                ListBox2.Items.Add("Pad Outer is empty");
            }
            else
            {
                ListBox2.Items.Remove("Pad Outer is empty");
            }

        }

        private void Underlayinner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(underlayinner.Text))
            {
                ListBox2.Items.Add("Underlay Inner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Underlay Inner is empty");
            }
        }

        private void Underlayouter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(underlayouter.Text))
            {
                ListBox2.Items.Add("Underlay Outer is empty");
            }
            else
            {
                ListBox2.Items.Remove("Underlay Outer is empty");
            }

        }

        private void Insulatorinner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(insulatorinner.Text))
            {
                ListBox2.Items.Add("Insulator Inner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Insulator Inner is empty");
            }

        }

        private void Insulatorouter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(insulatorouter.Text))
            {
                ListBox2.Items.Add("Insulator Outer is empty");
            }
            else
            {
                ListBox2.Items.Remove("Insulator Outer is empty");
            }
        }

        private void Chamferinner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(chamferinner.Text))
            {
                ListBox2.Items.Add("Chamfer Inner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Chamfer Inner is empty");
            }

        }

        private void Chamferouter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(chamferouter.Text))
            {
                ListBox2.Items.Add("Chamfer Outer is empty");
            }
            else
            {
                ListBox2.Items.Remove("Chamfer Outer is empty");
            }
        }

        private void Slotinner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(slotinner.Text))
            {
                ListBox2.Items.Add("Slot Inner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Slot Inner is empty");
            }
        }

        private void Slotouter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(slotouter.Text))
            {
                ListBox2.Items.Add("Slot Outer is empty");
            }
            else
            {
                ListBox2.Items.Remove("Slot Outer is empty");
            }
        }

        private void Compinner_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(compinner.Text))
            {
                ListBox2.Items.Add("Compression Inner is empty");
            }
            else
            {
                ListBox2.Items.Remove("Compression Inner is empty");
            }
        }

        private void Compouter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(compouter.Text))
            {
                ListBox2.Items.Add("Compression Outer is empty");
            }
            else
            {
                ListBox2.Items.Remove("Compression Outer is empty");
            }
        }

        private void ListBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Button_Click_11(object sender, RoutedEventArgs e)
        {
            if (ListBox2.Items.Count < 1)
            {
                Button7.IsEnabled = true;
                feedback.Text = "Form Validated Header tab enabled.";
            }
            else
            {
                Button7.IsEnabled = false;
                feedback.Text = "Some fields still require attention.";
            }
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            Testlist.Items.Clear();
            string path = @"K:\Dyno Testing\" + olddyno.SelectedValue;

            string[] dirs = Directory.GetDirectories(path);
            // For folders in the directory
            foreach (string dir in dirs)
                Testlist.Items.Add(Path.GetFileName(dir));

        }

        private void Button_Click_12(object sender, RoutedEventArgs e)
        {
            TabControl1.SelectedIndex = 10;


        }

        private void Reassign_Click(object sender, RoutedEventArgs e)
        {
            if (olddyno.SelectedValue == newdyno.SelectedValue)
            {
                feedback.Text = "You cannot Reassign to the same Dyno.";
            }
            else
            {
                Directory.Move(@"K:\\Dyno Testing\\" + olddyno.SelectedValue + "\\" + Testlist.SelectedValue, @"K:\\Dyno Testing\\" + newdyno.SelectedValue + "\\" + Testlist.SelectedValue);
                feedback.Text = "Test: " + Testlist.SelectedValue + " Has been Reassigned from Dyno: " + olddyno.SelectedValue + " to Dyno: " + newdyno.SelectedValue;
                Testlistnew.Items.Clear();
                string path = @"K:\Dyno Testing\" + newdyno.SelectedValue;

                string[] dirs = Directory.GetDirectories(path);
                // For folders in the directory
                foreach (string dir in dirs)
                    Testlistnew.Items.Add(Path.GetFileName(dir));
                Testlist.Items.Remove(Testlist.SelectedValue);
            }
        }

        private void Testlist_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Directory.Exists("K:\\Dyno Testing\\" + olddyno.SelectedValue + "\\" + Testlist.SelectedValue))
            {
                reassign.IsEnabled = true;

            }
            else
            {
                reassign.IsEnabled = false;
            }
        }

        private void Newdyno_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Testlistnew.Items.Clear();
            string path = @"K:\Dyno Testing\" + newdyno.SelectedValue;

            string[] dirs = Directory.GetDirectories(path);
            // For folders in the directory
            foreach (string dir in dirs)
                Testlistnew.Items.Add(Path.GetFileName(dir));
        }

        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        private void Feedback_TextChanged(object sender, TextChangedEventArgs e)
        {

        }


        private void ComboBox1_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {
            if (Selecttest.Items.Count <= 0)
            {
                feedback.Text = "No Tests Found Please save Fox TR in excel form to K:\\NewTestRequests\\";
            }
            else
            {
                feedback.Text = "Importing Data........";
                pBar1.Value = 30;
                MessageBox.Show("Are you sure you want to Import Test: " + Selecttest.SelectedValue, "Selected Test: " + Selecttest.SelectedValue);

                {
                    //do something
                    Header.Text = "";
                    date.Text = DateTime.Now.ToShortDateString();
                    time.Text = DateTime.Now.ToString("hh: mm :ss tt");
                    Button1.IsEnabled = true;
                    Button2.IsEnabled = true;
                    Button3.IsEnabled = true;
                    Button4.IsEnabled = true;
                    Button5.IsEnabled = true;
                    Button6.IsEnabled = true;
                    Button7.IsEnabled = false;
                    Createfolder.IsEnabled = true;
                    Button8.IsEnabled = false;
                    Button9.IsEnabled = true;
                    MakeHeader.IsEnabled = true;
                    Excel excel = new Excel(@"K:\NewTestRequests\" + Selecttest.SelectedValue + ".xlsx", 1);
                    TestLog.Text = excel.ReadCell(14, 23) + " " + excel.ReadCell(16, 3) + " " + excel.ReadCell(14, 3) + " " + excel.ReadCell(24, 3) + " " + excel.ReadCell(14, 3) + " " + excel.ReadCell(24, 3) + " " + excel.ReadCell(25, 3) + " " + excel.ReadCell(26, 3) + " " + excel.ReadCell(23, 3) + " " + excel.ReadCell(42, 4);
                    TScript.Text = excel.ReadCell(42, 4);
                    TSpec.Text = excel.ReadCell(40, 4);
                    string testnumber = excel.ReadCell(0, 0);
                    testnumber = testnumber.Replace("Test Card-# (LV).: ", "")
                    .Replace("Test Card-# (NVH).: ", "");
                    pBar1.Value = 30;
                    Testnum.Text = testnumber;
                    Dyno.Text = excel.ReadCell(48, 4);
                    Submittedby.Text = excel.ReadCell(6, 2);
                    owner.Text = excel.ReadCell(7, 2);
                    projectname.Text = excel.ReadCell(8, 2);
                    projectnumber.Text = excel.ReadCell(6, 25);
                    reason.Text = excel.ReadCell(9, 2);
                    Brakename.Text = excel.ReadCell(14, 23);
                    rotorid.Text = excel.ReadCell(36, 3);
                    rotorsize.Text = excel.ReadCell(33, 3);
                    rotorsource.Text = excel.ReadCell(35, 3);
                    batchinner.Text = excel.ReadCell(14, 3);
                    batchouter.Text = excel.ReadCell(14, 8);
                    padinner.Text = excel.ReadCell(19, 3);
                    padouter.Text = excel.ReadCell(19, 8);
                    pBar1.Value = 40;
                    underlayinner.Text = excel.ReadCell(23, 3);
                    underlayouter.Text = excel.ReadCell(23, 8);
                    chamferinner.Text = excel.ReadCell(25, 3);
                    chamferouter.Text = excel.ReadCell(25, 8);
                    slotinner.Text = excel.ReadCell(26, 3);
                    slotouter.Text = excel.ReadCell(26, 8);
                    insulatorinner.Text = excel.ReadCell(24, 3);
                    insulatorouter.Text = excel.ReadCell(24, 8);
                    compinner.Text = excel.ReadCell(20, 3);
                    compouter.Text = excel.ReadCell(20, 8);
                    pistonnum.Text = excel.ReadCell(15, 23);
                    reqinertia.Text = excel.ReadCell(23, 23);
                    pBar1.Value = 50;
                    pistondia.Text = excel.ReadCell(16, 23);
                    rollingrad.Text = excel.ReadCell(21, 23);
                    effradius.Text = excel.ReadCell(22, 23);
                    if (excel.ReadCell(17, 3) == null)
                    {
                        Materialsalesinner.Text = excel.ReadCell(16, 3);
                        Materialsalesouter.Text = Materialsalesinner.Text;

                    }
                    else
                    {
                        Materialsalesinner.Text = excel.ReadCell(17, 3);
                        Materialsalesouter.Text = Materialsalesinner.Text;
                    }



                    excel.Close();

                }
                ListBox2.Items.Clear();
                if (String.IsNullOrEmpty(pistonnum.Text))
                {
                    ListBox2.Items.Add("Piston Number is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Piston Number is empty");
                }
                if (String.IsNullOrEmpty(pistondia.Text))
                {
                    ListBox2.Items.Add("Piston Diameter is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Piston Diameter is empty");
                }

                if (String.IsNullOrEmpty(rollingrad.Text))
                {
                    ListBox2.Items.Add("Rolling Radius is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Rolling Radius is empty");
                }
                if (String.IsNullOrEmpty(effradius.Text))
                {
                    ListBox2.Items.Add("Effective Radius by is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Effective Radius by is empty");
                }
                if (String.IsNullOrEmpty(reqinertia.Text))
                {
                    ListBox2.Items.Add("Required Inertia is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Required Inertia is empty");
                }
                if (String.IsNullOrEmpty(Materialsalesinner.Text))
                {
                    ListBox2.Items.Add("Material Sales Inner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Material Sales Inner is empty");
                }
                if (String.IsNullOrEmpty(Materialsalesouter.Text))
                {
                    ListBox2.Items.Add("Material Sales Outer is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Material Sales Outer is empty");
                }
                if (String.IsNullOrEmpty(TestLog.Text))
                {
                    ListBox2.Items.Add("Test Log is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Test Log is empty");
                }
                if (String.IsNullOrEmpty(TSpec.Text))
                {
                    ListBox2.Items.Add("Test Specification is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Test Specification is empty");
                }

                if (String.IsNullOrEmpty(Testnum.Text))
                {
                    ListBox2.Items.Add("Test Number is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Test Number is empty");
                }
                if (String.IsNullOrEmpty(Submittedby.Text))
                {
                    ListBox2.Items.Add("Submitted by is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Submitted by is empty");
                }
                if (String.IsNullOrEmpty(owner.Text))
                {
                    ListBox2.Items.Add("Owner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Owner is empty");
                }
                if (String.IsNullOrEmpty(projectname.Text))
                {
                    ListBox2.Items.Add("Project Name is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Project Name is empty");
                }
                if (String.IsNullOrEmpty(projectnumber.Text))
                {
                    ListBox2.Items.Add("Project Number is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Project Number is empty");
                }
                if (String.IsNullOrEmpty(reason.Text))
                {
                    ListBox2.Items.Add("Reason is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Reason is empty");
                }
                if (String.IsNullOrEmpty(Brakename.Text))
                {
                    ListBox2.Items.Add("Brake name is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Brake name is empty");
                }
                if (String.IsNullOrEmpty(rotorid.Text))
                {
                    ListBox2.Items.Add("Rotor ID is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Rotor ID is empty");
                }
                if (String.IsNullOrEmpty(rotorsize.Text))
                {
                    ListBox2.Items.Add("Rotor Size is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Rotor Size is empty");
                }
                if (String.IsNullOrEmpty(rotorsource.Text))
                {
                    ListBox2.Items.Add("Rotor Source is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Rotor Source is empty");
                }
                if (String.IsNullOrEmpty(batchinner.Text))
                {
                    ListBox2.Items.Add("Batch Inner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Batch Inner is empty");
                }
                if (String.IsNullOrEmpty(batchouter.Text))
                {
                    ListBox2.Items.Add("Batch Outer is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Batch Outer is empty");
                }
                if (String.IsNullOrEmpty(padinner.Text))
                {
                    ListBox2.Items.Add("Pad Inner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Pad Inner is empty");
                }
                if (String.IsNullOrEmpty(padouter.Text))
                {
                    ListBox2.Items.Add("Pad Outer is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Pad Outer is empty");
                }
                if (String.IsNullOrEmpty(underlayinner.Text))
                {
                    ListBox2.Items.Add("Underlay Inner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Underlay Inner is empty");
                }
                if (String.IsNullOrEmpty(underlayouter.Text))
                {
                    ListBox2.Items.Add("Underlay Outer is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Underlay Outer is empty");
                }
                if (String.IsNullOrEmpty(chamferinner.Text))
                {
                    ListBox2.Items.Add("Chamfer Inner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Chamfer Inner is empty");
                }
                if (String.IsNullOrEmpty(chamferouter.Text))
                {
                    ListBox2.Items.Add("Chamfer Outer is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Chamger Outer is empty");
                }
                if (String.IsNullOrEmpty(slotinner.Text))
                {
                    ListBox2.Items.Add("Slot Inner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Slot Inner is empty");
                }
                if (String.IsNullOrEmpty(slotouter.Text))
                {
                    ListBox2.Items.Add("Slot Outer is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Slot Outer is empty");
                }
                if (String.IsNullOrEmpty(insulatorinner.Text))
                {
                    ListBox2.Items.Add("Insulator Inner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Insulator Inner is empty");
                }
                if (String.IsNullOrEmpty(insulatorouter.Text))
                {
                    ListBox2.Items.Add("Insulator Outer is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Insulator Outer is empty");
                }
                if (String.IsNullOrEmpty(compinner.Text))
                {
                    ListBox2.Items.Add("Compression Inner is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Compression Inner is empty");
                }
                if (String.IsNullOrEmpty(compouter.Text))
                {
                    ListBox2.Items.Add("Compression Outer is empty");
                }
                else
                {
                    ListBox2.Items.Remove("Compression Outer is empty");
                }
                pBar1.Value = 100;
                if (Directory.Exists(@"k:\Dyno Testing\" + Dyno.Text + "\\" + Selecttest.SelectedValue))
                {
                    Createfolder.IsEnabled = false;
                    feedback.Text = "Test Folder Already Created, Please Create Script.";
                    Button8.IsEnabled = true;
                }
                else
                {
                    Createfolder.IsEnabled = true;
                    feedback.Text = "Create Test Folder";
                }
            }

        }

        private void Button_Click_13(object sender, RoutedEventArgs e)
        {
        }
    }
}



















