using System;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.IO;
using System.Runtime.InteropServices;

/*
 * A simple speech recognizer for predefined keywords
 * 
 * Author: Mohsen Parisay - Apr.17.2018
 * 
 */
namespace SpeechRecognitionSample
{
    public partial class Form1 : Form
    {

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        //Mouse actions
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        private string[] KEYWORDS = new string[] { "select", "stop" };
        private String outputPath = "C:\\TEMP\\EyeTAP\\Commands\\voice_commands.txt";

        public Form1()
        {
            InitializeComponent();            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            createSpeechRecognizer();            
        }

        private void createSpeechRecognizer()
        {
            SpeechRecognizer recognizer = new SpeechRecognizer();
            Choices commands = new Choices();
            commands.Add(KEYWORDS);

            GrammarBuilder gb = new GrammarBuilder();
            gb.Append(commands);

            Grammar commandsGrammer = new Grammar(gb);
            recognizer.LoadGrammar(commandsGrammer);
            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
        }

        public void DoLeftMouseClick()
        {
            //Call the imported function with the cursor's current position
            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
        }

        public void DoRightMouseClick()
        {
            //Call the imported function with the cursor's current position
            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, X, Y, 0, 0);
        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs arguments)
        {
            string result = arguments.Result.Text;
            ShowResult(result);
            
            if (result.Equals("select"))
            {
                DoLeftMouseClick();
            }            

            WriteToFile(result + "\n");
        }

        void ShowResult(string text)
        {
            resultLabel.Text = text;
        }

        void WriteToFile(String text)
        {                       
            // write with OVERWRITE option
            File.WriteAllText(outputPath, text);           
        }

    }
}
