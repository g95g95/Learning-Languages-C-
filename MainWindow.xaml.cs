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
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace NewWords
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public int attempt = 0;
        public Excel.Worksheet xlsSheet;
        public string[] fremd;
        public string[] known;
        //public int totalWords = fremd.Count();
        public int guessed = 0;
        public MainWindow()
        {
            
            InitializeComponent();
        }
        
        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            fileExcel <string> xls = new fileExcel<string>(filePath.Text, false);
            Excel.Worksheet xlsSheet = xls.returnWorksheet();
            fremd = fileExcel<string>.getColumnsbyHeader("Fremd", xlsSheet);
            known = fileExcel<string>.getColumnsbyHeader("Translation", xlsSheet);
            Debug.WriteLine(String.Join("\n",fremd));
            Debug.WriteLine(String.Join("\n", known));
            System.Random rd = new System.Random();
            int random = rd.Next(0, fremd.Count());
            if (directionTranslation.Text == "0")
            {
                deutscheWort.Text = fremd[random];
                SpeechSynthesizer synthesizer = new SpeechSynthesizer();
                synthesizer.Volume = 100;  // 0...100
                synthesizer.Rate = -2;     // -10...10

                // Synchronous
                synthesizer.Speak("Hello World");

            }
            else if (directionTranslation.Text == "1")
            {
                knownWord.Text = known[random];
            }
            else 
            {
                System.Windows.Application.Current.Shutdown();
            }

        }


        private void browseButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                filePath.Text = openFileDialog.FileName;

            }
        }

        private void resetButton_Click(object sender, RoutedEventArgs e)
        {
            //filePath.Text = "";
            deutscheWort.Text = "";
            knownWord.Text = "";
            guessedWords.Text = "";
            successRateo.Text = "";
            remainingWords.Text = "";
            attemptN.Text = "";
        }

        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void confirmationButton_Click(object sender, RoutedEventArgs e)
        {


            if (directionTranslation.Text == "0")
            {
                try
                {
                    System.Random rd = new System.Random();
                    int random;
                    string rightWord = known[Array.IndexOf(fremd, deutscheWort.Text)];
                    //Debug.WriteLine(string.Join(" ",(rightWord.Split("/"))));
                    attempt += 1;
                    List<string> barWords = new List<string>(rightWord.Split(" / "));
                    barWords = barWords.ConvertAll(d => d.ToLower());
                    Debug.WriteLine(barWords[0]);

                    if (knownWord.Text.ToLower() == rightWord.ToLower())
                    {
                        fremd = fremd.Where(val => val != deutscheWort.Text).ToArray();
                        known = known.Where(val => val != rightWord).ToArray();
                        //Debug.WriteLine("Sono qui " + fremd.Count().ToString() + "     " + known.Count().ToString());
                        guessed += 1;
                        conclusion.Text = "Right! Bravo!\n";
                        random = rd.Next(0, fremd.Count());
                        deutscheWort.Text = fremd[random];
                    }

                    else if (barWords.Contains(knownWord.Text.ToLower()))
                    {
                        fremd = fremd.Where(val => val != deutscheWort.Text).ToArray();
                        known = known.Where(val => val != rightWord).ToArray();
                        guessed += 1;
                        random = rd.Next(0, fremd.Count());
                        deutscheWort.Text = fremd[random];
                        conclusion.Text = "It's all good man, nobody is perfect";
                    }

                    attemptN.Text = attempt.ToString();
                    remainingWords.Text = Convert.ToString(fremd.Count());
                    guessedWords.Text = guessed.ToString();
                    successRateo.Text = Convert.ToString(Convert.ToDouble(guessed) / Convert.ToDouble(attempt));
                    random = rd.Next(0, fremd.Count());
                    deutscheWort.Text = fremd[random];
                    knownWord.Text = "";
                    correctAnswer.Text = "The right answer was: " + rightWord;
                }

                catch (IndexOutOfRangeException IOOREx)
                {
                    var confirm2 = (MessageBox.Show("Do you want to keep your score?\n", "Save or not!", MessageBoxButton.YesNo));
                    var confirm1 = (MessageBox.Show("You ended the list of words. Do you want to restart?\n", "Remain or Leave!", MessageBoxButton.YesNo));
                    Debug.WriteLine(Convert.ToString(confirm2) + " " + Convert.ToString(confirm1));
                    resetButton_Click(sender, e);
                    if (confirm1 == MessageBoxResult.No)
                    {
                        System.Windows.Application.Current.Shutdown();
                    }
                    if (confirm2 == MessageBoxResult.No)
                    {
                        attempt = 0;
                        guessed = 0;
                    }
                }


            }


        else if (directionTranslation.Text == "1")
            {
                try 
                {
                    System.Random rd = new System.Random();
                    int random;
                    string rightWord = fremd[Array.IndexOf(known, knownWord.Text)];
                    attempt += 1;
                    List<string> barWords = new List<string>(rightWord.Split(" / "));
                    barWords = barWords.ConvertAll(d => d.ToLower());

                    if (deutscheWort.Text.ToLower() == rightWord.ToLower())
                    {
                        fremd = fremd.Where(val => val != rightWord).ToArray();
                        known = known.Where(val => val != knownWord.Text).ToArray();
                        //Debug.WriteLine("Sono qui " + fremd.Count().ToString() + "     " + known.Count().ToString());
                        guessed += 1;
                        conclusion.Text = "Right! Bravo!\n";
                        random = rd.Next(0, known.Count());
                        knownWord.Text = known[random];
                    }

                    else if (barWords.Contains(knownWord.Text.ToLower()))
                    {
                        fremd = fremd.Where(val => val != knownWord.Text).ToArray();
                        known = known.Where(val => val != rightWord).ToArray();
                        guessed += 1;
                        random = rd.Next(0, fremd.Count());
                        knownWord.Text = known[random];
                        conclusion.Text = "It's all good man, nobody is perfect";
                    }

                    attemptN.Text = attempt.ToString();
                    remainingWords.Text = Convert.ToString(fremd.Count());
                    guessedWords.Text = guessed.ToString();
                    successRateo.Text = Convert.ToString(Convert.ToDouble(guessed) / Convert.ToDouble(attempt));
                    random = rd.Next(0, known.Count());
                    knownWord.Text = known[random];
                    deutscheWort.Text = "";
                    correctAnswer.Text = "The right answer was: " + rightWord;
                }

                catch (IndexOutOfRangeException IOOREx)
                {
                    var confirm2 = (MessageBox.Show("Do you want to keep your score?\n", "Save or not!", MessageBoxButton.YesNo));
                    var confirm1 = (MessageBox.Show("You ended the list of words. Do you want to restart?\n", "Remain or Leave!", MessageBoxButton.YesNo));
                    Debug.WriteLine(Convert.ToString(confirm2) + " " + Convert.ToString(confirm1));
                    resetButton_Click(sender, e);
                    if (confirm1 == MessageBoxResult.No)
                    {
                        System.Windows.Application.Current.Shutdown();
                    }
                    if (confirm2 == MessageBoxResult.No)
                    {
                        attempt = 0;
                        guessed = 0;
                    }
                }

            }
        }
    }
}
