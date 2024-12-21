using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace SpectrometerColorViewer
{
    public partial class Form1 : Form
    {
        private double[] x10 = { 0.0191, 0.0434, 0.0847, 0.1406, 0.2045, 0.2647, 0.3147, 0.3577, 0.3837, 0.3867, 0.3707, 0.3430, 0.3023,
                                 0.2541, 0.1956, 0.1323, 0.0805, 0.0411, 0.0162, 0.0051, 0.0038, 0.0154, 0.0375, 0.0714, 0.1177, 0.1730,
                                 0.2365, 0.3042, 0.3768, 0.4516, 0.5298, 0.6161, 0.7052, 0.7938, 0.8787, 0.9512, 1.0142, 1.0743, 1.1185,
                                 1.1343, 1.1240, 1.0891, 1.0305, 0.9507, 0.8563, 0.7549, 0.6475, 0.5351, 0.4316, 0.3437, 0.2683, 0.2043,
                                 0.1526, 0.1122, 0.0813, 0.0579, 0.0409, 0.0286, 0.0199, 0.0138, 0.0096};

        private double[] y10 = { 0.0020, 0.0045, 0.0088, 0.0145, 0.0214, 0.0295, 0.0387, 0.0496, 0.0621, 0.0747, 0.0895, 0.1063, 0.1282,
                                 0.1528, 0.1852, 0.2199, 0.2536, 0.2977, 0.3391, 0.3954, 0.4608, 0.5314, 0.6067, 0.6857, 0.7618, 0.8233,
                                 0.8752, 0.9238, 0.9620, 0.9822, 0.9918, 0.9991, 0.9973, 0.9824, 0.9556, 0.9152, 0.8689, 0.8256, 0.7774,
                                 0.7204, 0.6583, 0.5939, 0.5280, 0.4618, 0.3981, 0.3396, 0.2835, 0.2283, 0.1798, 0.1402, 0.1076, 0.0812,
                                 0.0603, 0.0441, 0.0318, 0.0226, 0.0159, 0.0111, 0.0077, 0.0054, 0.0037 };

        private double[] z10 = { 0.0860, 0.1971, 0.3894, 0.6568, 0.9725, 1.2825, 1.5535, 1.7985, 1.9673, 2.0273, 1.9948, 1.9007, 1.7454,
                                 1.5549, 1.3176, 1.0302, 0.7721, 0.5701, 0.4153, 0.3024, 0.2185, 0.1592, 0.1120, 0.0822, 0.0607, 0.0431,
                                 0.0305, 0.0206, 0.0137, 0.0079, 0.0040, 0.0011, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
                                 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
                                 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000 };

        private double[] d65 = { 82.75, 87.12, 91.49, 92.46, 93.43, 90.06, 86.68, 95.77, 104.86, 110.94, 117.01, 117.41, 117.81, 116.34,
                                 114.86, 115.39, 115.92, 112.37, 108.81, 109.08, 109.35, 108.58, 107.80, 106.30, 104.79, 106.24, 107.69,
                                 106.05, 104.41, 104.23, 104.05, 104.02, 100.00, 98.17, 96.33, 96.06, 95.79, 92.24, 88.69, 89.35, 90.01,
                                 89.80, 89.60, 88.65, 87.70, 85.49, 83.29, 83.49, 83.70, 81.86, 80.03, 80.12, 80.21, 81.25, 82.28, 80.28,
                                 78.28, 74.00, 69.72, 70.67, 71.61 };

        private double[] standartPlate = { 0.85800, 0.86810, 0.87820, 0.88565, 0.89310, 0.89670, 0.90030, 0.90605, 0.91180, 0.91350, 0.91520, 0.91745, 
                                           0.91970, 0.91955, 0.91940, 0.92260, 0.92580, 0.92235, 0.91890, 0.92265, 0.92640, 0.92855, 0.93070, 0.93495, 
                                           0.93920, 0.93810, 0.93700, 0.93515, 0.93330, 0.93470, 0.93610, 0.93900, 0.94190, 0.94155, 0.94120, 0.94090, 
                                           0.94060, 0.94165, 0.94270, 0.94590, 0.94910, 0.94875, 0.94840, 0.95125, 0.95410, 0.95655, 0.95900, 0.95940, 
                                           0.95980, 0.95595, 0.95210, 0.95440, 0.95670, 0.96190, 0.96710, 0.96710, 0.96710, 0.96650, 0.96590, 0.96810, 
                                           0.97030 };

        private double[] whiteFiber = { 0.94781904216722, 0.969005451702626, 0.990429626977382, 0.99659601626591, 0.98902299431596,
                                        0.977675171708042, 0.981346693644441, 0.982366177500768, 0.982897148331433, 0.977851355190612,
                                        0.976339494582408, 0.985625792102401, 0.977801548768473, 0.980035279898027, 0.973454310539197,
                                        0.974354354717541, 0.962592785597437, 0.964535356835331, 0.960669260577959, 0.952754971090742,
                                        0.939853637940181, 0.933936879323295, 0.92253271328811, 0.923126088452934, 0.916272674571566,
                                        0.89915351309767, 0.897510526143235, 0.895764049938789, 0.889407202853783, 0.882733278843594,
                                        0.88809851156674, 0.878369718824126, 0.874383647704641, 0.876877025534908, 0.879564669296107,
                                        0.878289054140554, 0.883436122408742, 0.878630957755799, 0.87362825727779, 0.876621797927713,
                                        0.872264560234674, 0.869072174734045, 0.868759207467664, 0.866726497321988, 0.867778292774711,
                                        0.856537602792915, 0.851341065914877, 0.858716529948473, 0.848444184093301, 0.84793072918571,
                                        0.848493557466452, 0.851509155210219, 0.834976951100049, 0.830514952967858, 0.823675241927502,
                                        0.82027379795378, 0.831751094797738, 0.826833754385965, 0.829629056572868, 0.82917911046539,
                                        0.818946843861761 };

        private double[] whitePlate = { 1.04810423578441, 1.02588545463996, 1.05009169415971, 1.01323751841008, 0.976436793238869,
                                        0.963613634217163, 0.960882926811596, 0.965691301763283, 0.960971609564455, 0.958574091556111, 
                                        0.957313029140588, 0.966349138838405, 0.966592539901478, 0.965600761675579, 0.972255775937287, 
                                        0.970401077760487, 0.966472670611888, 0.971024820573855, 0.977603115883589, 0.968463569114064, 
                                        0.962510547530226, 0.970043612979451, 0.958640235959906, 0.961194286697642, 0.950327629245337, 
                                        0.944207149179352, 0.942810807476146, 0.943689883981618, 0.938771789451794, 0.941415689176902, 
                                        0.943716217347122, 0.94103674059766, 0.939642289289931, 0.946248072549674, 0.953235918224614, 
                                        0.960345700605001, 0.961089392883811, 0.965680614921531, 0.962576952136094, 0.968616875030403, 
                                        0.972618987719385, 0.971777430901243, 0.976553449657593, 0.972145300830893, 0.974965949573507, 
                                        0.970871120720332, 0.963915002442957, 0.96557866961266, 0.975016611754188, 0.966838348420052, 
                                        0.965137022966587, 0.975148332702216, 0.960864402712009, 0.964989813328475, 0.979071463448677, 
                                        0.981329561121479, 0.980982439955127, 0.976942362573099, 0.963707064022367, 0.968097171333289,
                                        0.959605674635222 };

        private double[] grayPlate = { 0.0766407193933372, 0.0684605198596376, 0.0967215881377149, 0.105269776653942, 0.107643666048025,
                                       0.113429627202483, 0.116217396737103, 0.120120402274335, 0.12239197798399, 0.122257121558427,
                                       0.125354522511609, 0.126729810720799, 0.124453201970443, 0.125220509163021, 0.120958455648606,
                                       0.120682083370365, 0.11749325252684, 0.113156343668171, 0.107392997999667, 0.112672697904038,
                                       0.113298624786663, 0.113357105664825, 0.115042901112407, 0.114388969196352, 0.113729814864752,
                                       0.111885242340121, 0.111220897294419, 0.11088761597771, 0.10454580337759, 0.0940702198466584,
                                       0.0966841174009206, 0.106494866471573, 0.108022291704254, 0.109511209078859, 0.109946450586058,
                                       0.108375814675485, 0.106117402500673, 0.10302488218163, 0.101934397775742, 0.104469105414214,
                                       0.102925615273704, 0.100777750595543, 0.0938177084563289, 0.0888632470844985, 0.0927263296537883, 
                                       0.0947544561361793, 0.0935977563253526, 0.0955203333994422, 0.0952804971711073, 0.0940398826614645, 
                                       0.0934115777115453, 0.0953994355648828, 0.0914979281029432, 0.0897561144194261, 0.0926835214437274, 
                                       0.0946572228875054, 0.0930054427605081, 0.0895238128654971, 0.0850520418018834, 0.0867975281123218, 
                                       0.0867035244688457};

        private double[] redCardBoard = { 0.0729717836957372, 0.0422343636080134, 0.0676033829201133, 0.0641704589950658, 0.0510919537293566,
                                          0.059648642869583, 0.053929993317856, 0.0520623867504133, 0.0511130341467297, 0.0488347856371774,
                                          0.0480709132512282, 0.0462423590749346, 0.0430344827586207, 0.0439755844594747, 0.0390647677003429, 
                                          0.0370549743894545, 0.0333305232699766, 0.0325079967515109, 0.0240155673899317, 0.0276912559841591, 
                                          0.0301013199265372, 0.029598504909928, 0.0300326524033302, 0.0302259249698847, 0.0317798903992752, 
                                          0.0301430078323629, 0.030790873323749, 0.0317049340443808, 0.0246103038381895, 0.0160529416724435, 
                                          0.0214426624424652, 0.0325915368775751, 0.0375564077955082, 0.0433462664259448, 0.0463558767858387, 
                                          0.0579048669347521, 0.0683891593139368, 0.1017703548681, 0.139451637733655, 0.217833341440872, 
                                          0.28096256152737, 0.375848440184873, 0.429498842752923, 0.48836752263963, 0.517516664576016, 
                                          0.543243217160566, 0.558102172688804, 0.570979418571983, 0.575508256116017, 0.580139493888711, 
                                          0.58070662964771, 0.592191931921259, 0.58370776454958, 0.58263936345445, 0.585537183382082, 
                                          0.59204708200062, 0.586268140989785, 0.580451134502924, 0.58438487890048, 0.584843009605132, 
                                          0.582014904630849 };

        private double[] blueCardBoard = { 0.359151907121736, 0.331403247179615, 0.35125234580543, 0.353756882737385, 0.357471250819131,
                                           0.357926538837071, 0.377692845218366, 0.398012868167177, 0.434265285681092, 0.459584395924173,
                                           0.498812648899657, 0.523904489160683, 0.532841659113301, 0.532544168744513, 0.528694628567568,
                                           0.519479204791735, 0.494292627494226, 0.471760190461815, 0.43687159202139, 0.40450392626273, 
                                           0.361609901881199, 0.321156369235466, 0.290707977965034, 0.246711899845121, 0.21574064122946, 
                                           0.17760741961946, 0.154738378345714, 0.128094319455309, 0.104968132141273, 0.0764387463074874, 
                                           0.0684013210592384, 0.0652455868434349, 0.0604697717923753, 0.0566246528469269, 0.0504794712313508, 
                                           0.0485128808247788, 0.0428675065020776, 0.0412393127289897, 0.0393406216194975, 0.041800150800214, 
                                           0.0417356237259484, 0.0394575972031794, 0.0342795448812632, 0.0293247889139528, 0.0353290077772203, 
                                           0.0407530257803162, 0.0431684729644214, 0.0461474001904559, 0.0486247813361474, 0.0521211924502047, 
                                           0.0560681194140535, 0.0586464921879904, 0.0588006397121214, 0.0576563585896344, 0.0577297005999293, 
                                           0.055025842819406, 0.0504991674878958, 0.0490893801169591, 0.0450611758799176, 0.0415703373478612, 
                                           0.0473190572381868 };

        private double[] lowGreen = { 0.0336948915989417, 0.0666666373165609, 0.0501081155433288, 0.0385373188689648, 0.0319760869565217,
                                      0.0342602714205097, 0.0308554848546907, 0.029320058384948, 0.0294846423699084, 0.032485614655666,
                                      0.0353195349244967, 0.0370350560184752, 0.0388652971814087, 0.0406556968228222, 0.0406247093344004,
                                      0.0444224636665674, 0.0505209029733587, 0.054536797370599, 0.066302728150146, 0.0852539397089397,
                                      0.114069082523121, 0.134831021568207, 0.143231341221291, 0.149064101692033, 0.146017807424594,
                                      0.139551564980362, 0.133601861845155, 0.127542336001675, 0.117968504370628, 0.0962823082023915,
                                      0.0932533834586466, 0.0871016693990934, 0.0813170405920312, 0.0689530114546853, 0.0605926448078475,
                                      0.0541137907505299, 0.0520382084038866, 0.0527051379690206, 0.0538016491746369, 0.0569563498764216,
                                      0.0559234753833209, 0.0543894889417329, 0.0469036985953178, 0.041423330924053, 0.0461779313780435,
                                      0.0500287605891369, 0.0488100618263217, 0.0529562218862038, 0.0560564751747046, 0.0574926790902391,
                                      0.0579886852085967, 0.0655163589137828, 0.0695778442898266, 0.0753107243509133, 0.0767368599528595,
                                      0.0767230169050715, 0.0754094299639754, 0.0753895196876497, 0.070386531279993, 0.0652317999566904,
                                      0.05361059734412 };

        private double[] mintGreen = { 0.0813323099921352, 0.113207613622881, 0.0962861072902338, 0.061438057885226, 0.046238768115942,
                                       0.047474462099603, 0.0392106582640844, 0.0369076773280176, 0.0384033586695251, 0.0400486307976181,
                                       0.0409639263761426, 0.0423671753581125, 0.0447719782882527, 0.0475297126151521, 0.0481433723838048,
                                       0.0537248387489853, 0.0614251994185653, 0.0672332682376261, 0.0888712648889433, 0.110781112266112,
                                       0.163225612464978, 0.214252027489845, 0.251132541818813, 0.290973803248023, 0.297859744779582,
                                       0.301249093640344, 0.298222623869366, 0.294129864195167, 0.285289340596512, 0.27260986620426,
                                       0.273180137351252, 0.280603867430693, 0.280539671979954, 0.28570310973726, 0.289188384049719,
                                       0.28989674222898, 0.290295233304777, 0.291517631360096, 0.291916678454647, 0.294207757206431,
                                       0.292544009932099, 0.296545641670057, 0.28885519221662, 0.287166140740567, 0.289939778317454,
                                       0.289234846045106, 0.291825621752901, 0.291849417801323, 0.294302697912191, 0.293851475264003,
                                       0.293660830172777, 0.295146620637983, 0.294132447942685, 0.292753314544962, 0.303050112306705,
                                       0.296568567713171, 0.29770228862047, 0.301289993642419, 0.291134014204136, 0.284230130097433,
                                       0.28584375723776 };

        private double w10x = 94.76;
        private double w10y = 99.98;
        private double w10z = 107.304;

        private List<double> deltaEList = new List<double>();
        private List<double> deltaLList = new List<double>();
        private List<double> deltaAList = new List<double>();
        private List<double> deltaBList = new List<double>();
        private double[] ERsLed = new double[61];

        private double lStandardSample = 0;
        private double aStandardSample = 0;
        private double bStandardSample = 0;
        private double deltaLamda = 5.0;

        public Thread thread;
        private bool threadState = true;
        private double firstPixelD = 0;
        private double lastPixelD = 0;
        private int firstNm = 400;
        private int lastNm = 700;

        private int testLoopCount = 3;
        private int otherLoopCount = 30;
        private int waitLoopCount = 5;

        private Random random = new Random();
        private double deltaGraphicsMinumum = 0;
        private double deltaGraphicsMaximum = 5;
        private int graphicMinimum = 0;
        private int graphicMaximum = 0;
        private int seriesIndex = 0;
        private int darkSeriesIndex = 0;
        private int ELedSeriesIndex = 0;
        private int standartSampleSeriesIndex = 0;

        private double trackbarValue = 0;
        private int firstTime = 0;

        public Form1()
        {
            InitializeComponent();

            groupBox7.Enabled = false;
            groupBox9.Enabled = false;
            groupBox4.Enabled = false;
            groupBox2.Enabled = false;
            groupBox6.Enabled = false;
            groupBox10.Enabled = false;
            groupBox8.Enabled = false;
            groupBox11.Enabled = false;
            groupBox12.Enabled = false;
            groupBox3.Enabled = false;
            groupBox5.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] serialPorts = SerialPort.GetPortNames();
            comboBox1.Items.Clear();
            foreach (string serialPort in serialPorts)
                comboBox1.Items.Add(serialPort);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string selectPortName = comboBox1.Text;
            Functions.OpenSerialPort(serialPort1, selectPortName, 3000000);

            if (!serialPort1.IsOpen)
            {
                label2.Text = "On";
                label2.ForeColor = Color.Green;

                groupBox7.Enabled = true;
                groupBox9.Enabled = true;
                groupBox4.Enabled = true;
                groupBox2.Enabled = true;
                groupBox6.Enabled = true;
                groupBox10.Enabled = true;
                groupBox8.Enabled = true;
                groupBox11.Enabled = true;
                groupBox12.Enabled = true;
                groupBox3.Enabled = true;
                groupBox5.Enabled = true;

                comboBox2.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;
                radioButton1.Checked = true;
                radioButton2.Checked = false;

                string[] pixelSettings = new string[2];
                string fileName = "PixelSettings.txt";
                pixelSettings = Functions.ReadFile(fileName);
                firstPixelD = Convert.ToDouble(pixelSettings[0]);
                label8.Text = pixelSettings[0];
                lastPixelD = Convert.ToDouble(pixelSettings[1]);
                label9.Text = pixelSettings[1];
            }
            else
            {
                label2.Text = "Off";
                label2.ForeColor = Color.Red;

                groupBox7.Enabled = false;
                groupBox9.Enabled = false;
                groupBox4.Enabled = false;
                groupBox2.Enabled = false;
                groupBox6.Enabled = false;
                groupBox10.Enabled = false;
                groupBox8.Enabled = false;
                groupBox11.Enabled = false;
                groupBox12.Enabled = false;
                groupBox3.Enabled = false;
                groupBox5.Enabled = false;

                radioButton1.Checked = false;
                radioButton2.Checked = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Functions.CloseSerialPort(serialPort1);

            groupBox7.Enabled = false;
            groupBox9.Enabled = false;
            groupBox4.Enabled = false;
            groupBox2.Enabled = false;
            groupBox6.Enabled = false;
            groupBox10.Enabled = false;
            groupBox8.Enabled = false;
            groupBox11.Enabled = false;
            groupBox12.Enabled = false;
            groupBox3.Enabled = false;
            groupBox5.Enabled = false;

            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            radioButton1.Checked = false;
            radioButton2.Checked = false;

            textBox30.Text = "";
            textBox5.Text = "10";
            textBox6.Text = "1";
            textBox3.Text = "0";
            textBox70.Text = "";
            textBox80.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox4.Text = "5";
            textBox20.Text = "3";
            textBox21.Text = "30";
            textBox22.Text = "5";

            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            listBox5.Items.Clear();
            listBox6.Items.Clear();

            chart5.Series.Clear();
            seriesIndex = 0;
            darkSeriesIndex = 0;
            ELedSeriesIndex = 0;
            standartSampleSeriesIndex = 0;
            chart1.Series.Clear();
            chart2.Series.Clear();
            chart3.Series.Clear();
            chart4.Series.Clear();

            label51.Text = "";
            label55.Text = "";
            label59.Text = "";
            label63.Text = "";
            label49.Text = "";
            label53.Text = "";
            label57.Text = "";
            label61.Text = "";

            label2.Text = "Off";
            label2.ForeColor = Color.Red;
            label8.Text = "X";
            label9.Text = "X";
            label24.Text = "Default";
            label24.ForeColor = Color.FromArgb(128, 64, 0);
            label19.Text = "Default";
            label19.ForeColor = Color.FromArgb(128, 64, 0);
            label23.Text = "Default";
            label23.ForeColor = Color.FromArgb(128, 64, 0);
            label12.Text = "Default";
            label12.ForeColor = Color.FromArgb(128, 64, 0);
            label5.Text = "Default";
            label5.ForeColor = Color.FromArgb(128, 64, 0);
            label28.Text = "Default";
            label28.ForeColor = Color.FromArgb(128, 64, 0);
            label25.Text = "Default";
            label25.ForeColor = Color.FromArgb(128, 64, 0);
            label36.Text = "Default";
            label36.ForeColor = Color.FromArgb(128, 64, 0);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            label5.Text = "Default";
            label5.ForeColor = Color.FromArgb(128, 64, 0);
            this.Refresh();
            Thread.Sleep(300);

            if (textBox1.Text != "" && textBox2.Text != "")
            {
                string fileName = "PixelSettings.txt";
                Functions.WriteFile(fileName, textBox1.Text, textBox2.Text);
                label5.Text = "Setting";
                label5.ForeColor = Color.Green;

                string[] pixelSettings = new string[2];
                pixelSettings = Functions.ReadFile(fileName);
                firstPixelD = Convert.ToDouble(pixelSettings[0]);
                label8.Text = pixelSettings[0];
                lastPixelD = Convert.ToDouble(pixelSettings[1]);
                label9.Text = pixelSettings[1];

                textBox1.Text = "";
                textBox2.Text = "";
            }
            else
            {
                label5.Text = "Error";
                label5.ForeColor = Color.Red;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                label12.Text = "Default";
                label12.ForeColor = Color.FromArgb(128, 64, 0);
                this.Refresh();
                Thread.Sleep(300);

                deltaGraphicsMinumum = Convert.ToDouble(textBox3.Text);
                deltaGraphicsMaximum = Convert.ToDouble(textBox4.Text);
                label12.Text = "True";
                label12.ForeColor = Color.Green;
            }
            catch (Exception)
            {
                label12.Text = "False";
                label12.ForeColor = Color.Red;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                label23.Text = "Default";
                label23.ForeColor = Color.FromArgb(128, 64, 0);
                this.Refresh();
                Thread.Sleep(300);

                testLoopCount = Convert.ToInt32(textBox20.Text);
                otherLoopCount = Convert.ToInt32(textBox21.Text);
                label23.Text = "True";
                label23.ForeColor = Color.Green;
            }
            catch (Exception)
            {
                label23.Text = "False";
                label23.ForeColor = Color.Red;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                label19.Text = "Default";
                label19.ForeColor = Color.FromArgb(128, 64, 0);
                label24.Text = "Default";
                label24.ForeColor = Color.FromArgb(128, 64, 0);
                this.Refresh();
                Thread.Sleep(300);

                string choosingIntegration = textBox5.Text;
                string averageValue = textBox6.Text;
                string sendData = "*MEASure:DARKspectra " + choosingIntegration + " " + averageValue + " format<CR>\r";

                int loopCount = otherLoopCount;
                int firstPixel = (int)firstPixelD - 1;
                int lastPixel = (int)lastPixelD;
                int dataCount = lastPixel - firstPixel;
                int section = (int)(((lastPixelD - firstPixelD) * (1.9868) / 5) + 1);
                double[] pixelSectionDataValue = new double[section];
                int choosingFilter = 10;
                int dataLenght = 2049;

                string choosingDigitalGain = comboBox2.SelectedIndex.ToString();
                string choosingAnalogGain = label18.Text;

                if (radioButton1.Checked == true)
                {
                    choosingFilter = 0;
                }
                else if (radioButton2.Checked == true)
                {
                    choosingFilter = 1;
                }

                if (Functions.DigitalGainSetting(serialPort1, choosingDigitalGain) == 6 && 
                    Functions.AnalogGainSetting(serialPort1, choosingAnalogGain) == 6)
                {
                    label19.Text = "True";
                    label19.ForeColor = Color.Green;
                }
                else
                {
                    label19.Text = "False";
                    label19.ForeColor = Color.Red;
                }

                pixelSectionDataValue = Functions.WaveCalculation(serialPort1, sendData, choosingFilter, section, loopCount, 
                                        dataLenght, dataCount, firstPixel, firstPixelD);
                label24.Text = "True";
                label24.ForeColor = Color.Green;

                double graphicsMinumumValue = 0;
                double graphicsMaximumValue = 0;
                seriesIndex = chart5.Series.Count;
                darkSeriesIndex = darkSeriesIndex + 1;
                chart5.Series.Add("Dark " + (darkSeriesIndex).ToString());
                chart5.ChartAreas["ChartArea1"].AxisX.Interval = 10;
                chart5.ChartAreas["ChartArea1"].AxisX.Minimum = firstNm;
                chart5.ChartAreas["ChartArea1"].AxisX.Maximum = lastNm;
                graphicsMinumumValue = pixelSectionDataValue[0];
                graphicsMaximumValue = pixelSectionDataValue[0];
                for (int i = 0; i < pixelSectionDataValue.Length; i++)
                {
                    if (graphicsMinumumValue > pixelSectionDataValue[i])
                        graphicsMinumumValue = pixelSectionDataValue[i];
                    if (graphicsMaximumValue < pixelSectionDataValue[i])
                        graphicsMaximumValue = pixelSectionDataValue[i];
                }

                if (graphicMaximum > graphicsMaximumValue)
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Maximum = graphicMaximum;
                }
                else
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Maximum = (int)graphicsMaximumValue + 1;
                    graphicMaximum = (int)graphicsMaximumValue + 1;
                }

                if (graphicMinimum > graphicsMinumumValue)
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Minimum = (int)graphicsMinumumValue + 1;
                    graphicMinimum = (int)graphicsMinumumValue + 1;
                }
                else
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Minimum = graphicMinimum;
                }

                for (int j = 10; j < 71; j++)
                {
                    chart5.Series[seriesIndex].ChartType = SeriesChartType.Line;
                    chart5.Series[seriesIndex].Color = Color.FromArgb(random.Next(0, 0), random.Next(0, 0), random.Next(0, 255));
                    chart5.Series[seriesIndex].Points.AddXY(350 + (j * 5), pixelSectionDataValue[j]);
                }
            }
            catch (Exception)
            {
                label24.Text = "False";
                label24.ForeColor = Color.Red;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                label25.Text = "Default";
                label25.ForeColor = Color.FromArgb(128, 64, 0);
                this.Refresh();
                Thread.Sleep(300);

                listBox1.Items.Clear();

                string choosingIntegration = textBox5.Text;
                string averageValue = textBox6.Text;
                string sendData = "*MEASure:REFERence " + choosingIntegration + " " + averageValue + " format<CR>\r";

                int loopCount = otherLoopCount;
                int firstPixel = (int)firstPixelD - 1;
                int lastPixel = (int)lastPixelD;
                int dataCount = lastPixel - firstPixel;
                int section = (int)(((lastPixelD - firstPixelD) * (1.9868) / 5) + 1);
                double[] ERefLed = new double[61];
                double[] RStandard = new double[61];
                double[] pixelSectionDataValue = new double[section];
                int choosingFilter = 10;
                int dataLenght = 2049;

                if (comboBox3.SelectedIndex == 0)
                {
                    for (int a = 0; a < RStandard.Length; a++)
                        RStandard[a] = standartPlate[a];
                }
                else if (comboBox3.SelectedIndex == 1)
                {
                    for (int b = 0; b < RStandard.Length; b++)
                        RStandard[b] = whiteFiber[b];
                }
                else if (comboBox3.SelectedIndex == 2)
                {
                    for (int c = 0; c < RStandard.Length; c++)
                        RStandard[c] = whitePlate[c];
                }
                else if (comboBox3.SelectedIndex == 3)
                {
                    for (int d = 0; d < RStandard.Length; d++)
                        RStandard[d] = grayPlate[d];
                }
                else if (comboBox3.SelectedIndex == 4)
                {
                    for (int f = 0; f < RStandard.Length; f++)
                        RStandard[f] = redCardBoard[f];
                }
                else if (comboBox3.SelectedIndex == 5)
                {
                    for (int g = 0; g < RStandard.Length; g++)
                        RStandard[g] = blueCardBoard[g];
                }
                else if (comboBox3.SelectedIndex == 6)
                {
                    for (int h = 0; h < RStandard.Length; h++)
                        RStandard[h] = lowGreen[h];
                }
                else if (comboBox3.SelectedIndex == 7)
                {
                    for (int k = 0; k < RStandard.Length; k++)
                        RStandard[k] = mintGreen[k];
                }
                else
                {
                    for (int p = 0; p < RStandard.Length; p++)
                        RStandard[p] = standartPlate[p];
                }

                if (radioButton1.Checked == true)
                    choosingFilter = 0;
                else if (radioButton2.Checked == true)
                    choosingFilter = 1;

                pixelSectionDataValue = Functions.WaveCalculation(serialPort1, sendData, choosingFilter, section, loopCount, 
                                        dataLenght, dataCount, firstPixel, firstPixelD);

                for (int i = 10; i < ERefLed.Length + 10; i++)
                    ERefLed[i - 10] = pixelSectionDataValue[i];
                for (int j = 0; j < ERsLed.Length; j++)
                {
                    ERsLed[j] = ERefLed[j] / RStandard[j];
                    listBox1.Items.Add(ERsLed[j]);
                }

                label25.Text = "True";
                label25.ForeColor = Color.Green;

                double graphicsMinumumValue = 0;
                double graphicsMaximumValue = 0;
                seriesIndex = chart5.Series.Count;
                ELedSeriesIndex = ELedSeriesIndex + 1;
                chart5.Series.Add("ELed " + (ELedSeriesIndex).ToString());
                chart5.ChartAreas["ChartArea1"].AxisX.Interval = 10;
                chart5.ChartAreas["ChartArea1"].AxisX.Minimum = firstNm;
                chart5.ChartAreas["ChartArea1"].AxisX.Maximum = lastNm;
                graphicsMinumumValue = pixelSectionDataValue[0];
                graphicsMaximumValue = pixelSectionDataValue[0];
                for (int r = 0; r < pixelSectionDataValue.Length; r++)
                {
                    if (graphicsMinumumValue > pixelSectionDataValue[r])
                        graphicsMinumumValue = pixelSectionDataValue[r];
                    if (graphicsMaximumValue < pixelSectionDataValue[r])
                        graphicsMaximumValue = pixelSectionDataValue[r];
                }

                if (graphicMaximum > graphicsMaximumValue)
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Maximum = graphicMaximum;
                }
                else
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Maximum = (int)graphicsMaximumValue + 1;
                    graphicMaximum = (int)graphicsMaximumValue + 1;
                }

                if (graphicMinimum > graphicsMinumumValue)
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Minimum = (int)graphicsMinumumValue + 1;
                    graphicMinimum = (int)graphicsMinumumValue + 1;
                }
                else
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Minimum = graphicMinimum;
                }

                for (int t = 10; t < 71; t++)
                {
                    chart5.Series[seriesIndex].ChartType = SeriesChartType.Line;
                    chart5.Series[seriesIndex].Color = Color.FromArgb(random.Next(0, 255), random.Next(0, 0), random.Next(0, 0));
                    chart5.Series[seriesIndex].Points.AddXY(350 + (t * 5), pixelSectionDataValue[t]);
                }
            }
            catch (Exception)
            {
                label25.Text = "False";
                label25.ForeColor = Color.Red;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                label28.Text = "Default";
                label28.ForeColor = Color.FromArgb(128, 64, 0);
                this.Refresh();
                Thread.Sleep(300);

                listBox2.Items.Clear();

                string choosingIntegration = textBox5.Text;
                string averageValue = textBox6.Text;
                string sendData = "*MEASure:REFERence " + choosingIntegration + " " + averageValue + " format<CR>\r";

                int loopCount = otherLoopCount;
                int firstPixel = (int)firstPixelD - 1;
                int lastPixel = (int)lastPixelD;
                int dataCount = lastPixel - firstPixel;
                int section = (int)(((lastPixelD - firstPixelD) * (1.9868) / 5) + 1);
                double[] pixelSectionDataValue = new double[section];
                double[] RStandardSample = new double[61];
                double[] EStandardSample = new double[61];
                int choosingFilter = 10;
                int dataLenght = 2049;

                double XStandardSample = 0;
                double XMultiStandardSample = 0;
                double YStandardSample = 0;
                double YMultiStandardSample = 0;
                double ZStandardSample = 0;
                double ZMultiStandardSample = 0;
                double kValue = KFunction();

                if (radioButton1.Checked == true)
                    choosingFilter = 0;
                else if (radioButton2.Checked == true)
                    choosingFilter = 1;

                pixelSectionDataValue = Functions.WaveCalculation(serialPort1, sendData, choosingFilter, section, loopCount, 
                                        dataLenght, dataCount, firstPixel, firstPixelD);

                for (int m = 10; m < EStandardSample.Length + 10; m++)
                    EStandardSample[m - 10] = pixelSectionDataValue[m];
                for (int n = 0; n < RStandardSample.Length; n++)
                {
                    RStandardSample[n] = EStandardSample[n] / ERsLed[n];
                    listBox2.Items.Add(RStandardSample[n]);
                }

                for (int i = 0; i < x10.Length; i++)
                    XMultiStandardSample = XMultiStandardSample + (RStandardSample[i] * d65[i] * x10[i] * deltaLamda);
                XStandardSample = XMultiStandardSample * kValue;
                textBox30.Text = Math.Round(XStandardSample, 4).ToString();

                for (int j = 0; j < y10.Length; j++)
                    YMultiStandardSample = YMultiStandardSample + (RStandardSample[j] * d65[j] * y10[j] * deltaLamda);
                YStandardSample = YMultiStandardSample * kValue;
                textBox70.Text = Math.Round(YStandardSample, 4).ToString();

                for (int k = 0; k < z10.Length; k++)
                    ZMultiStandardSample = ZMultiStandardSample + (RStandardSample[k] * d65[k] * z10[k] * deltaLamda);
                ZStandardSample = ZMultiStandardSample * kValue;
                textBox80.Text = Math.Round(ZStandardSample, 4).ToString();

                XStandardSample = XStandardSample / w10x;
                YStandardSample = YStandardSample / w10y;
                ZStandardSample = ZStandardSample / w10z;
                lStandardSample = (116 * Math.Pow(YStandardSample, 1.0 / 3.0)) - 16;
                aStandardSample = 500 * (Math.Pow(XStandardSample, 1.0 / 3.0) - Math.Pow(YStandardSample, 1.0 / 3.0));
                bStandardSample = 200 * (Math.Pow(YStandardSample, 1.0 / 3.0) - Math.Pow(ZStandardSample, 1.0 / 3.0));
                textBox11.Text = Math.Round(lStandardSample, 4).ToString();
                textBox10.Text = Math.Round(aStandardSample, 4).ToString();
                textBox9.Text = Math.Round(bStandardSample, 4).ToString();

                label28.Text = "True";
                label28.ForeColor = Color.Green;

                double graphicsMinumumValue = 0;
                double graphicsMaximumValue = 0;
                seriesIndex = chart5.Series.Count;
                standartSampleSeriesIndex = standartSampleSeriesIndex + 1;
                chart5.Series.Add("STNu " + (standartSampleSeriesIndex).ToString());
                chart5.ChartAreas["ChartArea1"].AxisX.Interval = 10;
                chart5.ChartAreas["ChartArea1"].AxisX.Minimum = firstNm;
                chart5.ChartAreas["ChartArea1"].AxisX.Maximum = lastNm;
                graphicsMinumumValue = pixelSectionDataValue[0];
                graphicsMaximumValue = pixelSectionDataValue[0];
                for (int a = 0; a < pixelSectionDataValue.Length; a++)
                {
                    if (graphicsMinumumValue > pixelSectionDataValue[a])
                        graphicsMinumumValue = pixelSectionDataValue[a];
                    if (graphicsMaximumValue < pixelSectionDataValue[a])
                        graphicsMaximumValue = pixelSectionDataValue[a];
                }

                if (graphicMaximum > graphicsMaximumValue)
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Maximum = graphicMaximum;
                }
                else
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Maximum = (int)graphicsMaximumValue + 1;
                    graphicMaximum = (int)graphicsMaximumValue + 1;
                }

                if (graphicMinimum > graphicsMinumumValue)
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Minimum = (int)graphicsMinumumValue + 1;
                    graphicMinimum = (int)graphicsMinumumValue + 1;
                }
                else
                {
                    chart5.ChartAreas["ChartArea1"].AxisY.Minimum = graphicMinimum;
                }

                for (int b = 10; b < 71; b++)
                {
                    chart5.Series[seriesIndex].ChartType = SeriesChartType.Line;
                    chart5.Series[seriesIndex].Color = Color.FromArgb(random.Next(0, 0), random.Next(0, 255), random.Next(0, 0));
                    chart5.Series[seriesIndex].Points.AddXY(350 + (b * 5), pixelSectionDataValue[b]);
                }
            }
            catch (Exception)
            {
                label28.Text = "False";
                label28.ForeColor = Color.Red;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                textBox12.Text = "";
                textBox13.Text = "";
                textBox14.Text = "";
                textBox15.Text = "";
                textBox16.Text = "";
                textBox17.Text = "";
                textBox18.Text = "";

                label51.Text = "";
                label55.Text = "";
                label59.Text = "";
                label63.Text = "";
                label49.Text = "";
                label53.Text = "";
                label57.Text = "";
                label61.Text = "";

                listBox3.Items.Clear();
                listBox4.Items.Clear();
                listBox5.Items.Clear();
                listBox6.Items.Clear();

                chart1.Series["Series1"].Points.Clear();
                chart2.Series["Series1"].Points.Clear();
                chart3.Series["Series1"].Points.Clear();
                chart4.Series["Series1"].Points.Clear();
                chart1.ChartAreas["ChartArea1"].AxisY.Minimum = deltaGraphicsMinumum;
                chart1.ChartAreas["ChartArea1"].AxisY.Maximum = deltaGraphicsMaximum;
                chart2.ChartAreas["ChartArea1"].AxisY.Minimum = deltaGraphicsMinumum;
                chart2.ChartAreas["ChartArea1"].AxisY.Maximum = deltaGraphicsMaximum;
                chart3.ChartAreas["ChartArea1"].AxisY.Minimum = deltaGraphicsMinumum;
                chart3.ChartAreas["ChartArea1"].AxisY.Maximum = deltaGraphicsMaximum;
                chart4.ChartAreas["ChartArea1"].AxisY.Minimum = deltaGraphicsMinumum;
                chart4.ChartAreas["ChartArea1"].AxisY.Maximum = deltaGraphicsMaximum;

                deltaEList.Clear();
                deltaLList.Clear();
                deltaAList.Clear();
                deltaBList.Clear();

                string choosingIntegration = textBox5.Text;
                string averageValue = textBox6.Text;
                string sendData = "*MEASure:REFERence " + choosingIntegration + " " + averageValue + " format<CR>\r";

                int loopCount = testLoopCount;
                int firstPixel = (int)firstPixelD - 1;
                int lastPixel = (int)lastPixelD;
                int dataCount = lastPixel - firstPixel;
                int section = (int)(((lastPixelD - firstPixelD) * (1.9868) / 5) + 1);
                int choosingFilter = 10;
                int dataLenght = 2049;
                firstTime = (DateTime.Now.Hour * 60 * 60) + (DateTime.Now.Minute * 60) + (DateTime.Now.Second);

                if (radioButton1.Checked == true)
                    choosingFilter = 0;
                else if (radioButton2.Checked == true)
                    choosingFilter = 1;

                label36.Text = "Start";
                label36.ForeColor = Color.Green;

                threadState = true;
                if (thread != null && thread.IsAlive == true)
                    return;
                thread = new Thread(() => Process(serialPort1, sendData, choosingFilter, section, loopCount, 
                                          dataLenght, dataCount, firstPixel, firstPixelD));
                thread.Start();
            }
            catch (Exception)
            {
                label36.Text = "Stop";
                label36.ForeColor = Color.Red;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            threadState = false;
            label36.Text = "Stop";
            label36.ForeColor = Color.Red;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            chart5.Series.Clear();
            seriesIndex = 0;
            darkSeriesIndex = 0;
            ELedSeriesIndex = 0;
            standartSampleSeriesIndex = 0;
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            trackbarValue = trackBar1.Value;
            trackbarValue = trackbarValue / 10;
            label18.Text = trackbarValue.ToString();
        }

        private void Process(SerialPort serialPort, string sendData, int choosingFilter, int section, int loopCount, int dataLenght, int dataCount, int firstPixel, double firstPixelD)
        {
            while (true)
            {
                if (threadState == true)
                {
                    double XTestSample = 0;
                    double XMultiTestSample = 0;
                    double YTestSample = 0;
                    double YMultiTestSample = 0;
                    double ZTestSample = 0;
                    double ZMultiTestSample = 0;

                    double kValue = KFunction();
                    double[] RTestSample = new double[61];
                    double[] ETestSample = new double[61];
                    double[] pixelSectionDataValue = new double[section];
                    int lastTime = 0;

                    double lDifference = 0;
                    double aDifference = 0;
                    double bDifference = 0;
                    double lDifferenceAbs = 0;
                    double aDifferenceAbs = 0;
                    double bDifferenceAbs = 0;

                    double deltaEResult = 0;
                    double deltaEAverage = 0;
                    double deltaLAverage = 0;
                    double deltaAAverage = 0;
                    double deltaBAverage = 0;

                    double lTestSample = 0;
                    double aTestSample = 0;
                    double bTestSample = 0;

                    pixelSectionDataValue = Functions.WaveCalculation(serialPort, sendData, choosingFilter, section, loopCount, 
                                            dataLenght, dataCount, firstPixel, firstPixelD);

                    for (int m = 10; m < ETestSample.Length + 10; m++)
                        ETestSample[m - 10] = pixelSectionDataValue[m];
                    for (int n = 0; n < RTestSample.Length; n++)
                        RTestSample[n] = ETestSample[n] / ERsLed[n];

                    for (int i = 0; i < x10.Length; i++)
                        XMultiTestSample = XMultiTestSample + (RTestSample[i] * d65[i] * x10[i] * deltaLamda);
                    XTestSample = XMultiTestSample * kValue;

                    for (int j = 0; j < y10.Length; j++)
                        YMultiTestSample = YMultiTestSample + (RTestSample[j] * d65[j] * y10[j] * deltaLamda);
                    YTestSample = YMultiTestSample * kValue;

                    for (int k = 0; k < z10.Length; k++)
                        ZMultiTestSample = ZMultiTestSample + (RTestSample[k] * d65[k] * z10[k] * deltaLamda);
                    ZTestSample = ZMultiTestSample * kValue;

                    XTestSample = XTestSample / w10x;
                    YTestSample = YTestSample / w10y;
                    ZTestSample = ZTestSample / w10z;
                    lTestSample = (116 * Math.Pow(YTestSample, 1.0 / 3.0)) - 16;
                    aTestSample = 500 * (Math.Pow(XTestSample, 1.0 / 3.0) - Math.Pow(YTestSample, 1.0 / 3.0));
                    bTestSample = 200 * (Math.Pow(YTestSample, 1.0 / 3.0) - Math.Pow(ZTestSample, 1.0 / 3.0));

                    lDifferenceAbs = lTestSample - lStandardSample;
                    lDifferenceAbs = Math.Abs(lDifferenceAbs);
                    lDifference = Math.Pow(lDifferenceAbs, 2);
                    aDifferenceAbs = aTestSample - aStandardSample;
                    aDifferenceAbs = Math.Abs(aDifferenceAbs);
                    aDifference = Math.Pow(aDifferenceAbs, 2);
                    bDifferenceAbs = bTestSample - bStandardSample;
                    bDifferenceAbs = Math.Abs(bDifferenceAbs);
                    bDifference = Math.Pow(bDifferenceAbs, 2);
                    deltaEResult = lDifference + aDifference + bDifference;
                    deltaEResult = Math.Sqrt(deltaEResult);

                    SetTextBox(textBox17, Math.Round(XTestSample, 5).ToString());
                    SetTextBox(textBox16, Math.Round(YTestSample, 5).ToString());
                    SetTextBox(textBox15, Math.Round(ZTestSample, 5).ToString());
                    SetTextBox(textBox14, Math.Round(lTestSample, 5).ToString());
                    SetTextBox(textBox13, Math.Round(aTestSample, 5).ToString());
                    SetTextBox(textBox12, Math.Round(bTestSample, 5).ToString());
                    SetTextBox(textBox18, Math.Round(deltaEResult, 5).ToString());
                    SetListBox(listBox3, Math.Round(deltaEResult, 5).ToString());
                    SetListBox(listBox4, Math.Round(lDifferenceAbs, 5).ToString());
                    SetListBox(listBox5, Math.Round(aDifferenceAbs, 5).ToString());
                    SetListBox(listBox6, Math.Round(bDifferenceAbs, 5).ToString());

                    if (deltaEList.Count > waitLoopCount)
                    {
                        deltaEList.Add(deltaEResult);
                        deltaLList.Add(lDifferenceAbs);
                        deltaAList.Add(aDifferenceAbs);
                        deltaBList.Add(bDifferenceAbs);

                        SetLabel(label49, Math.Round(deltaEResult, 5).ToString());
                        SetLabel(label53, Math.Round(lDifferenceAbs, 5).ToString());
                        SetLabel(label57, Math.Round(aDifferenceAbs, 5).ToString());
                        SetLabel(label61, Math.Round(bDifferenceAbs, 5).ToString());

                        lastTime = (DateTime.Now.Hour * 60 * 60) + (DateTime.Now.Minute * 60) + (DateTime.Now.Second);
                        lastTime = lastTime - firstTime;

                        SetChart(chart1, lastTime, Math.Round(deltaEResult, 5));
                        SetChart(chart2, lastTime, Math.Round(lDifferenceAbs, 5));
                        SetChart(chart3, lastTime, Math.Round(aDifferenceAbs, 5));
                        SetChart(chart4, lastTime, Math.Round(bDifferenceAbs, 5));

                        for (int a = waitLoopCount; a < deltaEList.Count; a++)
                        {
                            deltaEAverage = deltaEAverage + deltaEList[a];
                            deltaLAverage = deltaLAverage + deltaLList[a];
                            deltaAAverage = deltaAAverage + deltaAList[a];
                            deltaBAverage = deltaBAverage + deltaBList[a];
                        }
                        deltaEAverage = deltaEAverage / (deltaEList.Count - waitLoopCount);
                        deltaLAverage = deltaLAverage / (deltaLList.Count - waitLoopCount);
                        deltaAAverage = deltaAAverage / (deltaAList.Count - waitLoopCount);
                        deltaBAverage = deltaBAverage / (deltaBList.Count - waitLoopCount);

                        SetLabel(label51, Math.Round(deltaEAverage, 5).ToString());
                        SetLabel(label55, Math.Round(deltaLAverage, 5).ToString());
                        SetLabel(label59, Math.Round(deltaAAverage, 5).ToString());
                        SetLabel(label63, Math.Round(deltaBAverage, 5).ToString());
                    }
                    else
                    {
                        deltaEList.Add(deltaEResult);
                        deltaLList.Add(lDifferenceAbs);
                        deltaAList.Add(aDifferenceAbs);
                        deltaBList.Add(bDifferenceAbs);
                    }
                }
                else
                {
                    break;
                }
            }
        }

        private double KFunction()
        {
            double sumK = 0;
            double sumResult = 0;
            for (int i = 0; i < d65.Length; i++)
                sumK = sumK + (d65[i] * y10[i] * deltaLamda);
            sumResult = 100 / sumK;

            return sumResult;
        }

        delegate void SetChartCallback(Chart chart, int time, double value);
        private void SetChart(Chart chart, int time, double value)
        {
            if (chart.InvokeRequired)
            {
                SetChartCallback d = new SetChartCallback(_SetChart);
                chart.Invoke(d, new object[] { chart, time, value });
            }
            else
            {
                _SetChart(chart, time, value);
            }
        }

        private void _SetChart(Chart chart, int time, double value)
        {
            chart.Series[0].Points.AddXY(time, value);
        }

        delegate void SetListBoxCallback(ListBox listBox, string text);
        private void SetListBox(ListBox listBox, string text)
        {
            if (listBox.InvokeRequired)
            {
                SetListBoxCallback d = new SetListBoxCallback(_SetListBox);
                listBox.Invoke(d, new object[] { listBox, text });
            }
            else
            {
                _SetListBox(listBox, text);
            }
        }

        private void _SetListBox(ListBox listBox, string text)
        {
            listBox.Items.Insert(0, text);
        }

        delegate void SetTextBoxCallback(TextBox textBox, string text);
        private void SetTextBox(TextBox textBox, string text)
        {
            if (textBox.InvokeRequired)
            {
                SetTextBoxCallback d = new SetTextBoxCallback(_SetTextBox);
                textBox.Invoke(d, new object[] { textBox, text });
            }
            else
            {
                _SetTextBox(textBox, text);
            }
        }

        private void _SetTextBox(TextBox textBox, string text)
        {
            textBox.Text = text;
        }

        delegate void SetLabelCallback(Label label, string text);
        private void SetLabel(Label label, string text)
        {
            if (label.InvokeRequired)
            {
                SetLabelCallback d = new SetLabelCallback(_SetLabel);
                label.Invoke(d, new object[] { label, text });
            }
            else
            {
                _SetLabel(label, text);
            }
        }

        private void _SetLabel(Label label, string text)
        {
            label.Text = text;
        }
    }
}
