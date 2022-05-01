using System.IO.Compression;
//using Ionic.Zip;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.ComponentModel;
using System.Windows.Threading;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
//using Windows.UI.Xaml.Controls;
//using ComboBoxItem = System.Windows.Controls.ComboBoxItem;

namespace AutoLOInsert
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private IWebDriver driver;
        private string DriverPathTemp = @"C:\Users\aymen.menjli\Documents\Aymen menjli\TestBed";
        private String FolderPathForLO = "";
        private String FolderPathForBatchZip = "";
        private string MainTitle = "";
        private string LOCode = "";
        private string LOVersion = "1.0";
        private string Type = "Leçon";
        private string URL = "index.html";
        private string URLHelp = "./help/evaluation.html";
        private string Description = "";
        private string LODisplay = "";
        private string LOPosition = ""; 
        private string ParentLO = "";
        private string SecondParentLO = "";
        private string FilesPathZip = "";
        private int PosCount = 0;
        private int PosCountSecond = 0;
        private string CurrentTP1 = "";
        private int QuizNumber = 1;
        private int NumberOfQuizToAdd = 1;
        private int NumberOfEvalToAdd = 1;
        private List<LOData> LODataList = new List<LOData>();
        public List<String> AddedQuiz { get; set; } = new List<string>();
        public List<String> AddedEval { get; set; } = new List<string>();
        int CurrentMode = 0;



        public MainWindow()
        {
            InitializeComponent();
            if (NumberOfInstance != null)
                NumberOfInstance.Visibility = Visibility.Collapsed;
            if (AddEval != null)
                AddEval.Visibility = Visibility.Collapsed;
            
            //FolderPathForLO = @"D:\Atlas Work Menjli\E-Learning\Integration\TestFolderZip";
        }

        private void MainClickButton(object sender, RoutedEventArgs e)
        {
            switch (CurrentMode)
            {
                case 0:
                    UpdateData();
                    break;
                case 1:
                    AddQuizAsEval();
                    break;
                case 2:
                    AddEvaluation();
                    //if (NumberOfInstance != null)
                    //    NumberOfInstance.Visibility = Visibility.Visible;
                    break;
                default:
                    break;
            }

            //UpdateData();
            
        }


        private void AddQuizAsEval()
        {
            FolderPathForLO = FolderPathLO.Text;
            bool DriverIsGood = false;
            try
            {
                if (DevTemp.IsChecked == true)
                {
                    driver = new ChromeDriver(DriverPathTemp);
                    DriverIsGood = true;
                }
                else
                {
                    driver = new ChromeDriver();
                    DriverIsGood = true;
                }

            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show("Using default Driver and it's not compatible, please Update ChromeDriver");
            }

            if (!DriverIsGood)
                return;

            driver.Navigate().GoToUrl("https://");
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(12);
            Thread.Sleep(3000);

            //used for connection is not private
            var fieldAdvanced = driver.FindElement(By.Id("details-button"));
            Thread.Sleep(500);
            Actions Tempactions = new Actions(driver);
            Tempactions.MoveToElement(fieldAdvanced);
            Tempactions.Click();
            Tempactions.Build().Perform();

            var fieldProceed = driver.FindElement(By.Id("proceed-link"));
            Thread.Sleep(500);
            Actions Tempactions2 = new Actions(driver);
            Tempactions2.MoveToElement(fieldProceed);
            Tempactions2.Click();
            Tempactions2.Build().Perform();
            //Finished connection is not private

            Thread.Sleep(3000);
            var field1 = driver.FindElement(By.XPath("//input[@id='inputUsername']"));
            //driver.FindElement(By.XPath("//input[@id='inputUsername']"));

            //Login in
            AddInteraction("//input[@id='inputUsername']", 0, "TEMP");
            AddInteraction("//input[@id='inputPassword']", 0, "TEMP");
            AddInteraction("//button[@class='btn btn-button-box primary-color btn-sm']", 0, "");
            //click on Contenue pedagogique
            AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/atel-menu-first/div[@id='menu-first']/p-panelmenu[2]/div[1]", 2000, "");
            //click on learning object
            AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/atel-menu-second/div[@id='menu-second']/p-panelmenu/div/div[4]/div[1]/a[1]", 0, "");
            int QuizNumber = 0;
            //string[] dirs = Directory.GetDirectories(FolderPathForLO);
            NumberOfQuizToAdd = Int16.Parse(NumberOfInstance.Text);
            //foreach (string dir in dirs)
            for (int i = 0; i < NumberOfQuizToAdd; i++)
            {
                //Click on liste
                Thread.Sleep(1000);
                AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[1]//a[1]", 1000, "");
                LOCode = GetCode("Quiz_");

                //click on Ajouter to Add new LO
                AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[2]//a[1]//span[2]", 1000, "");
                //Add Title
                QuizNumber++;
                MainTitle = CompetenceCode.Text + " Quiz" + " " + QuizNumber;
                AddInteraction("//input[@id='titleInput']", 100, MainTitle);
                //Add LO Code
                AddInteraction("//input[@id='codeInput']", 0, LOCode);
                //Add LO Version
                AddInteraction("//input[@id='versionInput']", 0, LOVersion);
                //Add LO Type
                AddInteraction("//p-dropdown[@id='typeInput']", 0, "");
                Type = "Evaluation";
                AddInteraction(GetXpathForLOType(), 0, "");
                //Add LO Help URL
                AddInteraction("//input[@id='urlHelpInput']", 0, URLHelp);
                //Click on Chemins
                AddInteraction("//app-learning-object-detail//li[2]//a[1]", 0, "");
                //Click on ajouter Chemins
                AddInteraction("//button[contains(text(),'Ajouter un chemin')]", 100, "");
                //Click Select type
                AddInteraction("//body//p-dialog//div//div//div//div//div//div[2]", 50, "");
                //Select type
                LODisplay = "Précédent";
                AddInteraction(GetXpathForLODisplay(), 50, "");
                //Select type
                AddInteraction("//input[@id='treePositionInput']", 0, "2");
                //Save Chemins
                AddInteraction("//body//p-dialog//button[1]", 0, "");
                //Click on Configuration
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/ul[1]/li[3]/a[1]", 0, "");
                //Click on voir les solutions
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/p-inputswitch[1]", 0, "");
                //Click on Possibilité de refaire la question
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[3]/div[2]/p-inputswitch[1]", 0, "");
                //Click on Voir la solution après la réponse
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[2]/p-inputswitch[1]", 0, "");
                //Click on Navigation autorisée
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[5]/div[2]/p-inputswitch[1]", 0, "");
                //Click Select type d'exercice
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/p-dropdown[1]", 50, "");
                //Select type
                AddInteraction(GetXpathForExerciceType("Entrainement"), 50, "");
                //Click on Player
                AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/ul[1]/li[4]/a[1]", 0, "");
                AddInteraction("//textarea[@id='loPlayer']", 0, "{\u0022navigation\u0022:{\u0022buttons\u0022:[{\u0022target\u0022:\u0022player\u0022,\u0022action\u0022:\u0022previousSection\u0022},{\u0022target\u0022:\u0022player\u0022,\u0022action\u0022:\u0022nextSection\u0022}]},\u0022scripts\u0022:[{\u0022path\u0022:\u0022players/atel3_v0/0.2/scripts/atel/src/modules/themes/raspberry/deeppink/navigation/navigation.js\u0022,\u0022functions\u0022:[{\u0022name\u0022:\u0022MODULE_THEME_RASPBERRY_NAVIGATION.display\u0022,\u0022_args\u0022:{\u0022enabled\u0022:false}}]}]}");
                //return;
                //Click on save
                AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/app-learning-object-detail/div/button[1]", 500, "");
                //Click on liste
                AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[1]//a[1]", 1000, "");
                //got to the last LO Button and click to Edit
                ClickOnLastLOBtn();
                //Click On asset
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[1]/div[1]/div[2]/div[2]/div[1]/p-dropdown[1]", 500, "");
                //Select type
                AddInteraction("//p-dropdownitem[1]", 500, "");
                //Click on save
                AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/app-learning-object-detail/div/button[1]", 500, "");

            }

            AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[2]/atel-menu-second[1]/div[1]/p-panelmenu[1]/div[1]/div[2]/div[1]/a[1]", 0, "");

            AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[2]/atel-menu-second[1]/div[1]/p-panelmenu[1]/div[1]/div[2]/div[2]/div[1]/p-panelmenusub[1]/ul[1]/li[1]/a[1]", 1000, "");
            AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-list[1]/list-table[1]/p-table[1]/div[1]/div[1]/div[1]/div[2]/input[1]", 2000, CompetenceCode.Text);

            ClickOnCompetenceByCode(CompetenceCode.Text);

            //Click on Donnees
            AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/ul[1]/li[5]/a[1]", 300, "");
            for (int i = 0; i < NumberOfQuizToAdd; i++)
            {
                //Click on Quiz Item
                AddInteraction(GetXpathForParentLO("LO_ComonQuiz"), 50, "");
                //Click on Add button
                AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[5]/div[1]/div[1]/div[5]/div[1]/div[3]/div[1]/p-picklist[1]/div[1]/div[3]/div[1]/button[1]", 500, "");
                //for testing
                MainTitle = CompetenceCode.Text/*"Gen-18"*/ + " Quiz" + " " + (i+1);//uncomment the code for the final version
                AddedQuiz.Add(MainTitle);
            }

            for (int k = 0; k < AddedQuiz.Count; k++)
            {
                //Click on Quiz Item
                AddInteraction(GetXpathForParentLO(AddedQuiz[k]), 50, "");
                //Click on Add button
                AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[5]/div[1]/div[1]/div[5]/div[1]/div[3]/div[1]/p-picklist[1]/div[1]/div[3]/div[1]/button[1]", 500, "");

            }
            //Click on save
            AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-detail[1]/div[2]/button[1]", 500, "");

            if (AddEval.IsChecked == true)
            {
                AddEvaluation();
            }
            else
            {
                Thread.Sleep(25000);
                ResetData();
            }
        }

        private void AddEvaluation()
        {
            //driver = new ChromeDriver(DriverPathTemp);

            bool DriverIsGood = false;
            try
            {
                if (DevTemp.IsChecked == true)
                {
                    driver = new ChromeDriver(DriverPathTemp);
                    DriverIsGood = true;
                }
                else
                {
                    driver = new ChromeDriver();
                    DriverIsGood = true;
                }

            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show("Using default Driver and it's not compatible, please Update ChromeDriver");
            }

            if (!DriverIsGood)
                return;

            driver.Navigate().GoToUrl("https://");
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(12);
            Thread.Sleep(3000);

            //used for connection is not private
            var fieldAdvanced = driver.FindElement(By.Id("details-button"));
            Thread.Sleep(500);
            Actions Tempactions = new Actions(driver);
            Tempactions.MoveToElement(fieldAdvanced);
            Tempactions.Click();
            Tempactions.Build().Perform();

            var fieldProceed = driver.FindElement(By.Id("proceed-link"));
            Thread.Sleep(500);
            Actions Tempactions2 = new Actions(driver);
            Tempactions2.MoveToElement(fieldProceed);
            Tempactions2.Click();
            Tempactions2.Build().Perform();
            //Finished connection is not private

            Thread.Sleep(3000);
            var field1 = driver.FindElement(By.XPath("//input[@id='inputUsername']"));
            //driver.FindElement(By.XPath("//input[@id='inputUsername']"));

            //Login in
            AddInteraction("//input[@id='inputUsername']", 0, "TEMP");
            AddInteraction("//input[@id='inputPassword']", 0, "TEMP");
            AddInteraction("//button[@class='btn btn-button-box primary-color btn-sm']", 0, "");
            //click on Contenue pedagogique
            AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/atel-menu-first/div[@id='menu-first']/p-panelmenu[2]/div[1]", 2000, "");
            //click on learning object
            AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/atel-menu-second/div[@id='menu-second']/p-panelmenu/div/div[4]/div[1]/a[1]", 0, "");
            //int QuizNumber = 0;
            //string[] dirs = Directory.GetDirectories(FolderPathForLO);
            if(ComboBoxSelectType.SelectedIndex == 2)
            {
                NumberOfEvalToAdd = Int16.Parse(NumberOfInstance.Text);
                if (NumberOfEvalToAdd == 0)
                    NumberOfEvalToAdd = 1;
            }
            
            //foreach (string dir in dirs)
            for (int i = 0; i < NumberOfEvalToAdd; i++)
            {
                //Click on liste
                Thread.Sleep(1000);
                AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[1]//a[1]", 1000, "");
                LOCode = GetCode("Evaluation_");

                //click on Ajouter to Add new LO
                AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[2]//a[1]//span[2]", 1000, "");
                //Add Title
                //QuizNumber++;
                MainTitle = CompetenceCode.Text + " Evaluation";
                AddInteraction("//input[@id='titleInput']", 100, MainTitle);
                //Add LO Code
                AddInteraction("//input[@id='codeInput']", 0, LOCode);
                //Add LO Version
                AddInteraction("//input[@id='versionInput']", 0, LOVersion);
                //Add LO Type
                AddInteraction("//p-dropdown[@id='typeInput']", 0, "");
                Type = "Evaluation";
                AddInteraction(GetXpathForLOType(), 0, "");
                //Add LO Help URL
                AddInteraction("//input[@id='urlHelpInput']", 0, URLHelp);
                //Click on Chemins
                AddInteraction("//app-learning-object-detail//li[2]//a[1]", 0, "");
                //Click on ajouter Chemins
                AddInteraction("//button[contains(text(),'Ajouter un chemin')]", 100, "");
                //Click Select type
                AddInteraction("//body//p-dialog//div//div//div//div//div//div[2]", 50, "");
                //Select type
                LODisplay = "Précédent";
                AddInteraction(GetXpathForLODisplay(), 50, "");
                //Select type
                AddInteraction("//input[@id='treePositionInput']", 0, "1");
                //Save Chemins
                AddInteraction("//body//p-dialog//button[1]", 0, "");
                //Click on Configuration
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/ul[1]/li[3]/a[1]", 0, "");
                //Click on voir les solutions
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/p-inputswitch[1]", 0, "");
                //Click on Navigation automatique
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[6]/div[2]/p-inputswitch[1]", 0, "");
                //Click on Voir la solution après la réponse
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[2]/p-inputswitch[1]", 0, "");
                //Click on Répondu si ouvert
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[4]/div[2]/p-inputswitch[1]", 0, "");
                //Click Select Méthode de notation
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[7]/p-dropdown[1]", 50, "");
                //Select type
                AddInteraction("//p-dropdownitem[2]//li[1]", 50, "");
                //Click Select type d'exercice
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/p-dropdown[1]", 50, "");
                //Select type
                AddInteraction(GetXpathForExerciceType("Evaluation"), 100, "");
                //Click Select Type de statistiques
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[3]/div[1]/div[1]/div[1]/div[5]/div[2]/div[1]/div[1]/div[1]/p-dropdown[1]", 50, "");
                //Select type
                AddInteraction("//p-dropdownitem[3]//li[1]", 50, "");
                //Click on Ajouter des statistique
                AddInteraction("//button[contains(text(),'Ajouter des statistiques')]", 50, "");
                //Add Resultat URL
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-dialog[2]/div[1]/div[1]/div[2]/div[1]/textarea[1]", 500, "/stats/resultats.html");
                //Click on Save URL
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-dialog[2]/div[1]/div[1]/div[3]/p-footer[1]/button[1]", 50, "");
                //Click on Player
                AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/ul[1]/li[4]/a[1]", 0, "");
                AddInteraction("//textarea[@id='loPlayer']", 0, "{\u0022navigation\u0022:{\u0022buttons\u0022:[{\u0022target\u0022:\u0022player\u0022,\u0022action\u0022:\u0022previousSection\u0022},{\u0022target\u0022:\u0022player\u0022,\u0022action\u0022:\u0022nextSection\u0022}]},\u0022scripts\u0022:[{\u0022path\u0022:\u0022players/atel3_v0/0.2/scripts/atel/src/modules/themes/raspberry/deeppink/navigation/navigation.js\u0022,\u0022functions\u0022:[{\u0022name\u0022:\u0022MODULE_THEME_RASPBERRY_NAVIGATION.display\u0022,\u0022_args\u0022:{\u0022enabled\u0022:false}}]}]}");
                //return;
                //Click on save
                AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/app-learning-object-detail/div/button[1]", 500, "");
                //Click on liste
                AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[1]//a[1]", 1000, "");
                //got to the last LO Button and click to Edit
                ClickOnLastLOBtn();
                //Click On asset
                AddInteraction("//app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[1]/div[1]/div[2]/div[2]/div[1]/p-dropdown[1]", 500, "");
                //Select type
                AddInteraction("//p-dropdownitem[1]", 500, "");
                //Click on save
                AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/app-learning-object-detail/div/button[1]", 500, "");

            }

            AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[2]/atel-menu-second[1]/div[1]/p-panelmenu[1]/div[1]/div[2]/div[1]/a[1]", 0, "");

            AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[2]/atel-menu-second[1]/div[1]/p-panelmenu[1]/div[1]/div[2]/div[2]/div[1]/p-panelmenusub[1]/ul[1]/li[1]/a[1]", 1000, "");
            AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-list[1]/list-table[1]/p-table[1]/div[1]/div[1]/div[1]/div[2]/input[1]", 2000, CompetenceCode.Text);

            ClickOnCompetenceByCode(CompetenceCode.Text);

            //Click on Donnees
            AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/ul[1]/li[5]/a[1]", 300, "");
            for (int i = 0; i < NumberOfEvalToAdd; i++)
            {
                //Click on Quiz Item
                AddInteraction(GetXpathForParentLO("LO_Comon_000001"), 50, "");
                //Click on Add button
                AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[5]/div[1]/div[1]/div[5]/div[1]/div[3]/div[1]/p-picklist[1]/div[1]/div[3]/div[1]/button[1]", 500, "");
                //for testing
                MainTitle = CompetenceCode.Text/*"Gen-18"*/ + " Evaluation";//uncomment the code for the final version
                AddedEval.Add(MainTitle);
            }

            for (int k = 0; k < AddedEval.Count; k++)
            {
                //Click on Quiz Item
                AddInteraction(GetXpathForParentLO(AddedEval[k]), 50, "");
                //Click on Add button
                AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/div[1]/p-tabpanel[5]/div[1]/div[1]/div[5]/div[1]/div[3]/div[1]/p-picklist[1]/div[1]/div[3]/div[1]/button[1]", 500, "");

            }
            //Click on save
            AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-teaching-detail[1]/div[2]/button[1]", 500, "");
            
            Thread.Sleep(25000);
            ResetData();
        }
        private void UpdateData()
        {
            FolderPathForLO = FolderPathLO.Text;
            if(FolderPathForLO == "")
            {
                System.Windows.MessageBox.Show("Add Folder Path");
                return;
            }
            //ChromeOptions Options = new ChromeOptions();
            //Options.AddArgument("--headless");
            //driver = new ChromeDriver();
            //driver = new ChromeDriver(DriverPathTemp);


            bool DriverIsGood = false;
            try
            {
                if(DevTemp.IsChecked == true)
                {
                    driver = new ChromeDriver(DriverPathTemp);
                    DriverIsGood = true;
                }
                else
                {
                    driver = new ChromeDriver();
                    DriverIsGood = true;
                }
                
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show("Using default Driver and it's not compatible, please Update ChromeDriver");
            }

            if (!DriverIsGood)
                return;


            //driver.Navigate().GoToUrl("https://");
            driver.Navigate().GoToUrl("https://");
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(12);
            Thread.Sleep(3000);
            
            ////used for connection is not private
            //var fieldAdvanced = driver.FindElement(By.Id("details-button"));
            //Thread.Sleep(500);
            //Actions Tempactions = new Actions(driver);
            //Tempactions.MoveToElement(fieldAdvanced);
            //Tempactions.Click();
            //Tempactions.Build().Perform();

            //var fieldProceed = driver.FindElement(By.Id("proceed-link"));
            //Thread.Sleep(500);
            //Actions Tempactions2 = new Actions(driver);
            //Tempactions2.MoveToElement(fieldProceed);
            //Tempactions2.Click();
            //Tempactions2.Build().Perform();
            ////Finished connection is not private

            //Thread.Sleep(3000);
            //var field1 = driver.FindElement(By.XPath("//input[@id='inputUsername']"));
            ////driver.FindElement(By.XPath("//input[@id='inputUsername']"));

            //Login in
            AddInteraction("//input[@id='inputUsername']", 0, "TEMP");
            AddInteraction("//input[@id='inputPassword']", 0, "TEMP");
            AddInteraction("//button[@class='btn btn-button-box primary-color btn-sm']", 0, "");
            //click on Contenue pedagogique
            AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/atel-menu-first/div[@id='menu-first']/p-panelmenu[2]/div[1]", 2000, "");
            //click on learning object
            //AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/atel-menu-second/div[@id='menu-second']/p-panelmenu/div/div[4]/div[1]/a[1]", 0, "");
            AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[2]/atel-menu-second[1]/div[1]/p-panelmenu[1]/div[1]/div[6]/div[1]/a[1]", 0, "");

            ////Click on liste
            //AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[1]//a[1]", 1000, "");

            //MessageBox.Show(LOCode);

            string[] dirs = Directory.GetDirectories(FolderPathForLO);
            AddLOProgressBar.Maximum = dirs.Count();
            //Progress Bar in Diffrent thread
            //BackgroundWorker workerLO = new BackgroundWorker();
            //workerLO.WorkerReportsProgress = true;
            //workerLO.DoWork += workerLO_DoWork;
            //workerLO.ProgressChanged += workerLO_ProgressChanged;
            //workerLO.RunWorkerAsync();

            foreach (string dir in dirs)
            {
                //Click on liste
                Thread.Sleep(1000);
                //AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[1]//a[1]", 1000, "");
                AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[2]/atel-menu-second[1]/div[1]/p-panelmenu[1]/div[1]/div[6]/div[2]/div[1]/p-panelmenusub[1]/ul[1]/li[1]/a[1]", 1000, "");
                //LOCode = GetCode("LO_");
                
                //Thread.Sleep(5000);
                var TempFilesPathZip = Directory.GetFiles(dir, "*.zip", SearchOption.AllDirectories);
                FilesPathZip = TempFilesPathZip[0];
                System.Diagnostics.Debug.WriteLine(FilesPathZip);
                var files = Directory.GetFiles(dir, "*.json", SearchOption.AllDirectories);
                
                if (files[0] != null)
                {
                    string JsonFile = files[0];
                    LOCode = /*System.IO.Path.GetDirectoryName(JsonFile)*/new DirectoryInfo(System.IO.Path.GetDirectoryName(JsonFile)).Name; ;
                    //System.Windows.MessageBox.Show(LOCode);
                    //return;
                    string JsonF = System.IO.File.ReadAllText(JsonFile);

                    JObject jObject = JObject.Parse(JsonF);
                    JArray componentsObj = (JArray)jObject.SelectToken("components");
                    foreach (JToken componentObj in componentsObj)
                    {
                        string displayName = (string)componentObj.SelectToken("type");
                        if (displayName == "ComponentTitle")
                        {
                            JObject displayNameData = (JObject)componentObj.SelectToken("data");
                            string LOTitle1 = (string)displayNameData.SelectToken("level1");
                            string LOTitle2 = (string)displayNameData.SelectToken("level2");
                            string LOTitle3 = (string)displayNameData.SelectToken("level3");
                            //if (LOTitle1 != null && LOTitle1.Contains("'"))
                            //    LOTitle1 = LOTitle1.Replace("'", "’");
                            //if (LOTitle2 != null && LOTitle2.Contains("'"))
                            //    LOTitle2.Replace("'", "’");
                            //if (LOTitle3 != null && LOTitle3.Contains("'"))
                            //    LOTitle3.Replace("'", "’");
                            string LONumber = "";
                            try
                            {
                                LONumber = (string)displayNameData.SelectToken("number");
                            }
                            catch (Exception)
                            {

                                throw;
                            }

                            if(LOTitle1 != null)
                            {
                                LOTitle1 = LOTitle1.Trim();
                                LOTitle1 = LOTitle1.Replace("\n", " ");
                                LOTitle1 = Regex.Replace(LOTitle1, @"\s+", " ");
                            }
                                
                            if (LOTitle2 != null)
                            {
                                LOTitle2 = LOTitle2.Trim();
                                LOTitle2 = LOTitle2.Replace("\n", " ");
                                LOTitle2 = Regex.Replace(LOTitle2, @"\s+", " ");
                            }
                                
                            if (LOTitle3 != null)
                            {
                                LOTitle3 = LOTitle3.Trim();
                                LOTitle3 = LOTitle3.Replace("\n", " ");
                                LOTitle3 = Regex.Replace(LOTitle3, @"\s+", " ");
                            }
                                

                            string MainTitleNoNumber = LOTitle3;
                            string AdditionalNumber = LONumber != "" ? LONumber : "";
                            MainTitle = AdditionalNumber + "." + " " + LOTitle3;
                            if (LOTitle3 == null || LOTitle3 == "")
                            {
                                AdditionalNumber = LONumber != "" ? LONumber : "";
                                MainTitleNoNumber = LOTitle2;
                                MainTitle = AdditionalNumber + "." + " " + LOTitle2;
                            }
                            if (LOTitle2 == null || LOTitle2 == "")
                            {
                                AdditionalNumber = LONumber != "" ? LONumber : "";
                                MainTitleNoNumber = LOTitle1;
                                MainTitle = AdditionalNumber + "." + " " + LOTitle1;
                            }

                            if (MainTitle.IndexOf("Introduction", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                MainTitle = "Introduction";
                                LONumber = "1";
                            }
                                

                                LOData CurrentLOData = new LOData
                            {
                                LOName = MainTitleNoNumber,
                                LOCode = LOCode
                            };
                            LODataList.Add(CurrentLOData);
                            Description = GetDescription(LOTitle1, LOTitle2, LOTitle3);
                            //System.Diagnostics.Debug.WriteLine(LOTitle1, LOTitle2, LOTitle3);
                            UpdatePathData(LOTitle1, LOTitle2, LOTitle3, LONumber, LOCode);
                            //MessageBox.Show(Description);
                        }

                    }

                    var dirName = new DirectoryInfo(dir).Name;
                    if (dirName.IndexOf("Resume", StringComparison.OrdinalIgnoreCase) >= 0 || dirName.IndexOf("Résumé", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        //string FolderName = new DirectoryInfo(System.IO.Path.GetDirectoryName(dir)).Name;
                        
                        MainTitle = "Résumé";
                        if (dirName.IndexOf("suite", StringComparison.OrdinalIgnoreCase) >= 0 || dirName.Any(c => char.IsDigit(c)))
                            MainTitle = MainTitle + " (suite)";
                        UpdatePathData(MainTitle, "", "", "1", LOCode);
                       
                        LOData CurrentLOData = new LOData
                        {
                            LOName = MainTitle,
                            LOCode = LOCode
                        };
                        LODataList.Add(CurrentLOData);
                    }
                        
                    if (dirName.IndexOf("objectif", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        LOData CurrentLOData = new LOData
                        {
                            LOName = MainTitle,
                            LOCode = LOCode
                        };
                        LODataList.Add(CurrentLOData);
                        MainTitle = "Objectifs";
                        UpdatePathData(MainTitle, "", "", "1", LOCode);
                    }
                        
                    //if (dir.IndexOf("quiz", StringComparison.OrdinalIgnoreCase) >= 0)
                    //{
                    //    //MainTitle = "Quiz" + QuizNumber;
                    //    string TitleToPros = new DirectoryInfo(dir).Name;
                    //    string[] StringArray = TitleToPros.Split('-');
                    //    if(StringArray.Length == 3)
                    //    {
                    //        MainTitle = StringArray[1] + StringArray[0] + StringArray[2];
                    //    }
                    //    else
                    //    {
                    //        MainTitle = StringArray[1] + StringArray[0];
                    //    }
                    //    if (MainTitle.IndexOf("Suite", StringComparison.OrdinalIgnoreCase) >= 0)
                    //    {
                    //        LODisplay = "Précédent";
                    //        LOPosition = "";
                    //        ParentLO = "";
                    //    }
                    //    else
                    //    {
                    //        LODisplay = "Normal";
                    //        LOPosition = "0";
                    //        ParentLO = "";
                    //    }
                    //    QuizNumber++;
                    //}
                }
                else
                {
                    return;
                }

                //click on Ajouter to Add new LO
                AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[2]/atel-menu-second[1]/div[1]/p-panelmenu[1]/div[1]/div[6]/div[2]/div[1]/p-panelmenusub[1]/ul[1]/li[2]/a[1]", 1000, "");
                //AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[2]/atel-menu-second[1]/div[1]/p-panelmenu[1]/div[1]/div[6]/div[1]/a[1]/span[2]", 1000, "");
                //Add Title
                AddInteraction("//input[@id='titleInput']", 100, MainTitle);
                //Add LO Code
                AddInteraction("//input[@id='codeInput']", 0, LOCode);
                //Add LO Version
                AddInteraction("//input[@id='versionInput']", 0, LOVersion);
                //Add LO Type
                AddInteraction("//p-dropdown[@id='typeInput']", 0, "");
                AddInteraction(GetXpathForLOType(), 0, "");
                //Add LO URL
                AddInteraction("//input[@id='urlInput']", 0, URL);
                //Add LO Description
                AddInteraction("//textarea[@id='descriptionInput']", 0, Description);
                //Click on Chemins
                AddInteraction("//app-learning-object-detail//li[2]//a[1]", 0, "");
                //Click on ajouter Chemins
                AddInteraction("//button[contains(text(),'Ajouter un chemin')]", 100, "");
                //Click Select type
                AddInteraction("//body//p-dialog//div//div//div//div//div//div[2]", 50, "");
                //Select type
                AddInteraction(GetXpathForLODisplay(), 50, "");
                //Select type
                AddInteraction("//input[@id='treePositionInput']", 0, LOPosition);
                //Save Chemins
                AddInteraction("//body//p-dialog//button[1]", 0, "");
                //Add parent lo
                if (ParentLO != "")
                {
                    AddInteraction(GetXpathForParentLO(ParentLO), 50, "");
                    AddInteraction("//body//div[@id='settings-content']//div//div[3]//div[1]//button[1]", 0, "");
                    if(SecondParentLO != "")
                    {
                        AddInteraction(GetXpathForParentLO(SecondParentLO), 50, "");
                        AddInteraction("//body//div[@id='settings-content']//div//div[3]//div[1]//button[1]", 0, "");
                    }
                }
                //Click on Player
                AddInteraction("//body[1]/atel-app[1]/ng-component[1]/div[1]/div[3]/app-learning-object-detail[1]/div[1]/div[1]/p-tabview[1]/div[1]/ul[1]/li[3]", 0, "");
                AddInteraction("//textarea[@id='loPlayer']", 0, "{\u0022navigation\u0022:{\u0022buttons\u0022: [{\u0022target\u0022: \u0022player\u0022, \u0022action\u0022: \u0022previousSection\u0022},{\u0022target\u0022: \u0022player\u0022, \u0022action\u0022: \u0022nextSection\u0022}]}, \u0022scripts\u0022: [{\u0022path\u0022: \u0022players/atel3_v0/0.2/scripts/atel/src/modules/themes/raspberry/deeppink/navigation/auto_run/navigation_enabled.js\u0022, \u0022functions\u0022: []}]}");
                //return;
                //Click on save
                AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/app-learning-object-detail/div/button[1]", 500, "");

                
                //Click on liste
                AddInteraction("//div[4]//div[2]//div[1]//p-panelmenusub[1]//ul[1]//li[1]//a[1]", 1000, "");
                //Click on search Bar to find the last LO we saved based on the Code stored
                AddInteraction("//body/atel-app[1]/ng-component[1]/div[1]/div[3]/app-learning-object-list[1]/list-table[1]/p-table[1]/div[1]/div[1]/div[1]/div[2]/input[1]", 4000, LOCode);
                //Click on Found LO By given code 
                ClickOnCompetenceByCode(LOCode);
                //got to the last LO Button and click to Edit
                //ClickOnLastLOBtn();

                /***********Uncomment This to Debug without Uploading files****************/
                /***********Uncomment This to Debug without Uploading files****************/
                /***********Uncomment This to Debug without Uploading files****************/
                //Click on Fichier
                AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/app-learning-object-detail/div/div/p-tabview/div/ul/li[5]", 1000, "");

                //IWebElement ele = GetIwebElmByExplicitWait("//p-button[2]//button[1]");
                //driver.FindElement(By.XPath("//input[@class='ng-star-inserted']")).SendKeys(@FilesPathZip);
                IWebElement ele = driver.FindElement(By.XPath("//p-button[2]//button[1]"));
                ele.Click();
                Thread.Sleep(1000);
                System.Diagnostics.Debug.WriteLine(@FilesPathZip);
                SendKeys.SendWait(@FilesPathZip);
                SendKeys.SendWait("{Enter}");

                //Click on Chemins
                AddInteraction("//app-learning-object-detail//li[2]//a[1]", 2000, "");
                //Click on Fichier
                AddInteraction("//body/atel-app/ng-component/div[@id='settings-content']/div/app-learning-object-detail/div/div/p-tabview/div/ul/li[5]", 500, "");

                //Right Click on Zip File
                AddInteraction("//html//body//atel-app//ng-component//div//div//app-learning-object-detail//div//div//p-tabview//div//div//p-tabpanel//div//file-explorer//p-treetable//div//div//div//p-checkbox", 100, "");
                Thread.Sleep(1000);
                IWebElement UpZipFile = driver.FindElement(By.XPath("//td[contains(text(),'.zip')]"));
                //IWebElement UpZipFile = GetIwebElmByExplicitWait("//td[contains(text(),'.zip')]");
                Actions RightClick = new Actions(driver);
                RightClick.ContextClick(UpZipFile).Build().Perform();

                //Click on decompresser
                AddInteraction("//body//p-contextmenusub//li[8]", 1000, "");
                //Thread.Sleep(2000);
                AddInteraction("//body//p-dialog//button[1]", 2000, "");

                //Right Click on Zip File to delete
                Thread.Sleep(1000);
                IWebElement UpZipFile2 = driver.FindElement(By.XPath("//td[contains(text(),'.zip')]"));
                //IWebElement UpZipFile2 = GetIwebElmByExplicitWait("//td[contains(text(),'.zip')]");
                Actions RightClick2 = new Actions(driver);
                RightClick2.ContextClick(UpZipFile2).Build().Perform();

                //Click on supprimer
                AddInteraction("//a[@id='remove']", 500, "");
                AddInteraction("//body//p-confirmdialog//button[1]", 500, "");
                System.Windows.Application.Current.Dispatcher.Invoke(DispatcherPriority.ApplicationIdle, (Action)(() =>
                {
                    AddLOProgressBar.Value++;
                }));
            }
            ResetData();
        }


        private void ResetData()
        {
            FolderPathForLO = "";
            FolderPathForBatchZip = "";
            MainTitle = "";
            LOCode = "";
            Description = "";
            LODisplay = "";
            LOPosition = "";
            ParentLO = "";
            SecondParentLO = "";
            FilesPathZip = "";
            PosCount = 0;
            PosCountSecond = 0;
            CurrentTP1 = "";
            QuizNumber = 0;
            AddLOProgressBar.Value = 0;
            FolderPathLO.Clear();
            CompetenceName.Clear();
            CompetenceCode.Clear();
            QuizNumber = 1;
            NumberOfQuizToAdd = 1;
            AddedQuiz.Clear();
            AddedEval.Clear();
            driver.Quit();
            System.Windows.MessageBox.Show("Done");
        }

        private void AddInteraction(string PathToAdd,  int TimeOut, string KeyToAdd)
        {
            //IWebElement XPathToAdd = GetIwebElmByExplicitWait(PathToAdd);
            Thread.Sleep(TimeOut);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(12);
            var XPathToAdd = driver.FindElement(By.XPath(PathToAdd));
            if (XPathToAdd != null)
            {
                Actions actions = new Actions(driver);
                actions.MoveToElement(XPathToAdd);
                actions.Click();
                if(KeyToAdd != null && KeyToAdd != "")
                {
                    //actions.KeyDown(XPathToAdd, OpenQA.Selenium.Keys.ArrowLeft).KeyUp(XPathToAdd, OpenQA.Selenium.Keys.ArrowLeft).Build();
                    //actions.SendKeys(KeyToAdd);

                    System.Windows.Clipboard.SetText(KeyToAdd);
                    //actions.SendKeys(OpenQA.Selenium.Keys.Control + "v");
                    actions.KeyDown(OpenQA.Selenium.Keys.Control).SendKeys("v").KeyUp(OpenQA.Selenium.Keys.Control);

                    //IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
                    //jsExecutor.ExecuteScript("arguments[0].setAttribute('value', '" + KeyToAdd + "')", PathToAdd);
                    //jsExecutor.ExecuteScript("arguments[0].setAttribute(arguments[1], arguments[2]);", PathToAdd, "value", KeyToAdd);
                    //jsExecutor.ExecuteScript("document.evaluate('" + PathToAdd + "', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.setAttribute('value', '" + KeyToAdd + "');");

                    System.Diagnostics.Debug.WriteLine(KeyToAdd);
                    //if (KeyToAdd == "Backspace")
                    //{
                    //    //actions.SendKeys("X");
                    //    //Thread.Sleep(1000);
                    //    actions.SendKeys(OpenQA.Selenium.Keys.Backspace);
                    //}
                    //else
                    //{

                    //}
                    //actions.SendKeys(OpenQA.Selenium.Keys.Enter);
                }
                actions.Build().Perform();
                //HelpLabel.Content = "you clicked 9th";
            }
        }

        private IWebElement GetIwebElmByExplicitWait(string PathIn)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement XPathToAddOut = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath(PathIn)));
            return XPathToAddOut;
        }

        private string GetXpathForLOType()
        {
            string ReturnPath = "";
            if(Type == "Leçon")
            {
                ReturnPath = "//p-dropdownitem[1]//li[1]";
            }
            else
            {
                ReturnPath = "//p-dropdownitem[2]//li[1]";
            }
            return ReturnPath;
        }

        private void ClickOnLastLOBtn()
        {
            AddInteraction("//body/atel-app/ng-component/div/div/app-learning-object-list/list-table/p-paginator/div/a[4]", 1000, "");
            Thread.Sleep(1000);
            ReadOnlyCollection<IWebElement> elements = driver.FindElements(By.TagName("tr"));
            //MessageBox.Show(elements.Count.ToString());
            //LOCode = "LO_000" + (elements.Count - 1);
            for (int i = 0; i < elements.Count; i++)
            {
                
                if (elements.Count > 1)
                {
                    if (i == elements.Count - 1)
                    {
                        IWebElement Btnelement = elements[i].FindElement(By.TagName("button"));
                        Actions actions = new Actions(driver);
                        actions.MoveToElement(Btnelement);
                        actions.Click();
                        actions.Build().Perform();
                        //MessageBox.Show(i.ToString());
                    }
                }
                else
                {
                    IWebElement BtnelementS = elements[i].FindElement(By.TagName("button"));
                    Actions actionsS = new Actions(driver);
                    actionsS.MoveToElement(BtnelementS);
                    actionsS.Click();
                    actionsS.Build().Perform();
                }
                
            }
        }

        private void ClickOnCompetenceByCode(string CCode)
        {
            Thread.Sleep(2000);

            ReadOnlyCollection<IWebElement> elements = driver.FindElements(By.TagName("tr"));
            IWebElement ElementContainCode;
            for (int i = 0; i < elements.Count; i++)
            {
                //try
                //{
                //    ElementContainCode = elements[i].FindElement(By.XPath("//td[contains(text(),'" + CCode + "')]"));
                //}
                //catch (Exception)
                //{

                //    throw;
                //}
                if (elements[i].Text.Contains(CCode))
                {
                    IWebElement CurentElement = elements[i];
                    int currI = i;
                    IWebElement Btnelement = CurentElement.FindElement(By.TagName("button"));
                    string logElem = CurentElement.GetAttribute("innerHTML");
                    System.Diagnostics.Debug.WriteLine(logElem, currI);
                    Actions actions = new Actions(driver);
                    actions.MoveToElement(Btnelement);
                    actions.Click();
                    actions.Build().Perform();
                    //System.Diagnostics.Debug.WriteLine(Btnelement);
                }
            }
        }

        private string GetCode(string TypePrefix)
        {
            AddInteraction("//body/atel-app/ng-component/div/div/app-learning-object-list/list-table/p-paginator/div/a[4]", 1000, "");
            Thread.Sleep(1000);
            ReadOnlyCollection<IWebElement> elements;
            string eleCode;
            try
            {
                elements = driver.FindElements(By.TagName("tr"));
                ReadOnlyCollection<IWebElement> elementsTd = elements[elements.Count-1].FindElements(By.TagName("td"));
                eleCode = elementsTd[2].Text;
                //System.Windows.MessageBox.Show(eleCode);

            }
            catch (Exception)
            {
                throw;
            }
            string[] StringArray = eleCode.Split('_');
            string NeweleCode = StringArray[StringArray.Length-1];
            //string NeweleCode = eleCode.Remove(0, 3);
            int NewCodeFromString = Int16.Parse(NeweleCode) + 1;
            string CompCode = Regex.Replace(CompetenceCode.Text, @"\s+", "");
            string NewCode = TypePrefix + CompCode + "_" + "000" + NewCodeFromString;
            return NewCode;
        }

        private string GetDescription(string T1, string T2, string T3)
        {
            string NewDescription = CompetenceName.Text + "\r\n" + T1 + "\r\n" + T2 + "\r\n" + T3;
            return NewDescription;
        }

        private void UpdatePathData(string Tp1, string Tp2, string Tp3, string TNumber, string CurrCode)
        {
            
            if (CurrentTP1 == "")
            {
                CurrentTP1 = Tp1;
            }
            //TODO: Add system for the precedent position, LO with "(suite)" should be tagged with LODisplay = "Précédent"
            if (Tp2 == null || Tp2 == "" || Tp3 == null || Tp3 == "")
            {
                LODisplay = "Normal";
                LOPosition = "0";
                ParentLO = "";
                if (Tp1.IndexOf("suite", StringComparison.OrdinalIgnoreCase) >= 0 == true)
                {
                    LODisplay = "Précédent";
                    ParentLO = GetCodeFromLoName(CurrentTP1);
                }
                System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
            }

            if (Tp2 != null && Tp2.Length > 0 && Tp3 == null && LODisplay != "Précédent" || Tp2 != "" && Tp3 == "" && LODisplay != "Précédent")
            {
                if (Tp2.IndexOf("suite", StringComparison.OrdinalIgnoreCase) >= 0/*.Contains("(suite)")*/ == true)
                {
                    LODisplay = "Précédent";
                    //PosCount++;
                    System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
                }
                else
                {
                    LODisplay = "Normal";
                    //PosCount++;
                    System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
                }
                SecondParentLO = "";
                //string[] NumberList = TNumber.Split('.');
                //int locPos = (Int16.Parse(NumberList[1]) - 1) + PosCount;
                //LOPosition = locPos.ToString();
                ParentLO = /*Tp1*/GetCodeFromLoName(Tp1);
                System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
            }
            else if(Tp3 != "" && LODisplay != "Précédent")
            {
                if(Tp3 != null && Tp3.Length > 0)
                {
                    if (Tp3.IndexOf("suite", StringComparison.OrdinalIgnoreCase) >= 0/*.Contains("(suite)")*/ == true)
                    {
                        LODisplay = "Précédent";
                        System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
                        //PosCount++;
                    }
                    else
                    {
                        LODisplay = "Normal";
                        System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
                        //PosCount++;
                    }
                    //string[] NumberList = TNumber.Split('.');
                    //int locPos = (Int16.Parse(NumberList[2]) - 1) + PosCount;
                    //LOPosition = locPos.ToString();
                    ParentLO = /*Tp1*/GetCodeFromLoName(Tp1);
                    SecondParentLO = /*Tp2*/GetCodeFromLoName(Tp2);
                    System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
                }
            }
            bool CheckObjRes = CurrentTP1.IndexOf("objectifs", StringComparison.OrdinalIgnoreCase) >= 0 || CurrentTP1.IndexOf("resume", StringComparison.OrdinalIgnoreCase) >= 0 || CurrentTP1.IndexOf("Résumé", StringComparison.OrdinalIgnoreCase) >= 0 ? true : false;
            bool CheckNumberLength = TNumber != null && CurrentTP1 == Tp1 && TNumber.Length > 1;
            bool CheckSuite = Tp1.IndexOf("suite", StringComparison.OrdinalIgnoreCase) >= 0;
            bool CheckContainsSuite = (TNumber != null && CheckSuite);
            if (CheckNumberLength || CurrentTP1 == Tp1 && CheckObjRes && LODisplay == "Précédent" || CheckContainsSuite)
            {
                PosCount++;
               // string[] NumberList = TNumber.Split('.');
                int locPos = /*(Int16.Parse(NumberList[1]) - 1) + */PosCount;
                LOPosition = locPos.ToString();
                System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
            }
            else
            {
                SecondParentLO = "";
                PosCount = 0;
                LOPosition = "0";
                CurrentTP1 = Tp1;
                System.Diagnostics.Debug.WriteLine(ParentLO, LOPosition, LODisplay);
            }
            
            //MessageBox.Show(LODisplay);
            //MessageBox.Show(LOPosition);
            //MessageBox.Show(ParentLO);
        }

        private string GetCodeFromLoName(string TitleIn)
        {
            string CodeFound = "";
            foreach (LOData LOD in LODataList)
            {
                if(LOD.LOName == TitleIn)
                {
                    CodeFound = LOD.LOCode;
                    return CodeFound;
                }
            }

            return CodeFound;
        }

        private string GetXpathForLODisplay()
        {
            string ReturnPath = "";
            switch (LODisplay)
            {
                case "Normal":
                    ReturnPath = "//p-dropdownitem[1]//li[1]";
                    break;
                case "Précédent":
                    ReturnPath = "//p-dropdownitem[3]//li[1]";
                    break;
                default:
                    ReturnPath = "//p-dropdownitem[1]//li[1]";
                    break;
            }
            return ReturnPath;
        }

        private string GetXpathForExerciceType(string EType)
        {
            string ReturnPath = "";
            switch (EType)
            {
                case "Entrainement":
                    ReturnPath = "//p-dropdownitem[2]//li[1]";
                    break;
                case "Evaluation":
                    ReturnPath = "//p-dropdownitem[3]//li[1]";
                    break;
                default:
                    ReturnPath = "//p-dropdownitem[2]//li[1]";
                    break;
            }
            return ReturnPath;
        }

        private string GetXpathForParentLO(string Plo)
        {
            string ploModified = Plo;
            if (ploModified.Contains("'"))
                ploModified = ploModified.Replace("'", "’");
            string testXpathDy = "//span[contains(text(),'" + ploModified.ToString() + "')]";
            return testXpathDy;
        }

        private void BatchZipFiles_Click(object sender, RoutedEventArgs e)
        {
            FolderPathForBatchZip = FolderPathZip.Text;
            if(FolderPathForBatchZip == "")
            {
                System.Windows.MessageBox.Show("Add Folder Path");
                return;
            }
            var result = System.Windows.MessageBox.Show("This will delete all fla files \ncontinue?", "caption", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                

                
                string[] dirs = Directory.GetDirectories(FolderPathForBatchZip);
                ZipFilesProgressBar.Maximum = dirs.Count();
                int FCount = 0;
                foreach (string dir in dirs)
                {
                    try
                    {
                        //delete All FLA Files
                        var Flafiles = Directory.GetFiles(dir, "*.fla", SearchOption.AllDirectories);
                        //System.Windows.MessageBox.Show(Flafiles[0]);
                        File.Delete(Flafiles[0]);
                    }
                    catch (Exception)
                    {

                    }
                    //Zip Files 
                    var files = Directory.GetFiles(dir);
                    string[] Subdirs = Directory.GetDirectories(dir, "*", SearchOption.AllDirectories);

                    string FileNameToSave = dir + ".zip";
                    FCount++;
                    //MessageBox.Show(dir);
                    ZipFile.CreateFromDirectory(dir, FileNameToSave);

                    var Zipfiles = Directory.GetFiles(FolderPathForBatchZip, "*.zip", SearchOption.AllDirectories);
                    string NewFilePath = dir + @"\" + System.IO.Path.GetFileName(Zipfiles[0]);
                    File.Move(Zipfiles[0], NewFilePath);
                    //ZipFilesProgressBar.Value++;
                }

                //Thread.Sleep(1000);
                //var Zipfiles = Directory.GetFiles(FolderPathForBatchZip, "*.zip", SearchOption.AllDirectories);

                //for (int Zf = 0; Zf < Zipfiles.Length; Zf++)
                //{
                //    //MessageBox.Show(Zipfiles[Zf]);
                //    string NewFilePath = dirs[Zf] + @"\" + System.IO.Path.GetFileName(Zipfiles[Zf]);
                //    //File.Move(Zipfiles[Zf], NewFilePath);
                //}

                //Progress Bar in Diffrent thread
                BackgroundWorker worker = new BackgroundWorker();
                worker.WorkerReportsProgress = true;
                worker.DoWork += worker_DoWork;
                worker.ProgressChanged += worker_ProgressChanged;

                worker.RunWorkerAsync();

                //System.Windows.Application.Current.Dispatcher.Invoke(DispatcherPriority.ApplicationIdle, (Action)(() =>
                //{
                //    ZipFilesProgressBar.Value++;
                //}));
            }
            System.Windows.MessageBox.Show("Done");
            ZipFilesProgressBar.Value = 0;
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int i = 0; i < 100; i++)
            {
                (sender as BackgroundWorker).ReportProgress(i);
                Thread.Sleep(100);
            }
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ZipFilesProgressBar.Value = e.ProgressPercentage;
        }

        //ProgressBar for LO
        void workerLO_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int k = 0; k < AddLOProgressBar.Maximum; k++)
            {
                (sender as BackgroundWorker).ReportProgress(k);
                Thread.Sleep(100);
            }
        }

        void workerLO_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            AddLOProgressBar.Value = e.ProgressPercentage;
        }

        private void SelectLOFolderEvent(object sender, MouseButtonEventArgs e)
        {
            //System.Windows.MessageBox.Show("You selected: ");
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            //dialog.InitialDirectory = "C:/Users/aymen.menjli/Documents/Aymen menjli/TestBed/TestAppPPTX";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                FolderPathLO.Text = dialog.FileName;
                string lastFolderName = System.IO.Path.GetFileName(dialog.FileName);
                CompetenceCode.Text = lastFolderName;
                CompetenceName.Text = lastFolderName;
                //System.Windows.MessageBox.Show("You selected: " + lastFolderName);
            }
        }

        private void SelectLOZipFolderEvent(object sender, MouseButtonEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            //dialog.InitialDirectory = "C:/Users/aymen.menjli/Documents/Aymen menjli/TestBed/TestAppPPTX";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                FolderPathZip.Text = dialog.FileName;
                //MessageBox.Show("You selected: " + dialog.FileName);
            }
        }

        private void ComboBoxSelectType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //ComboBoxItem SelectedValue = (ComboBoxItem)ComboBoxSelectType.SelectedItem;
            CurrentMode = ComboBoxSelectType.SelectedIndex;
            switch (CurrentMode)
            {
                case 0:
                    if (NumberOfInstance != null)
                    {
                        NumberOfInstance.Visibility = Visibility.Collapsed;
                    }
                    //Thread.Sleep(10000);
                    if(AddEval != null)
                        AddEval.Visibility = Visibility.Collapsed;
                    break;
                case 1:
                    if (NumberOfInstance != null)
                        NumberOfInstance.Visibility = Visibility.Visible;
                    AddEval.Visibility = Visibility.Visible;
                    break;
                case 2:
                    if (NumberOfInstance != null)
                        NumberOfInstance.Visibility = Visibility.Visible;
                    AddEval.Visibility = Visibility.Collapsed;
                    break;
                default:
                    break;
            }
        }
    }
}
