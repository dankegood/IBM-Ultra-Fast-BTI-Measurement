using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Timers;
using System.Threading.Tasks;
using System.Collections;

namespace WGFMU_BTI_AK
{
    class Program
    {
        // No need
        static int[,] DUT_channels = new int[10, 4];        // 10 DTUS and 4 channels per DUT
                                                             // values sotres in this array will be assigned to channel, 2, 3 & 4
        static double[] ListvGateStress = new double[10];    //10 DTUS

        // Need for array
        //These are only default values and user is to provide all new values from VB GUI interface
        //static int channel1 = 901; //Taking WGFMU1's channel #1 as common gate channel --> This must be user input coming from VB
        //static int channel2 = 902; //Taking WGFMU1's channel #2 as drain channel --> This must be user input coming from VB
        //static int channel3 = 801; //Taking WGFMU2's channel #1 as drain channel --> This must be user input coming from VB, not used for individual devie stress
        //static int channel4 = 802; //Taking WGFMU2's channel #2 as drain channel --> This must be user input coming from VB, not used for individual devie stress
        //static int channel5 = 701;
        //static int channel6 = 702;
        //static int channel7 = 601;
        //static int channel8 = 602;
        //static int channel9 = 501;
        //static int channel10 = 502;
        static int[] channel1 = new int[5];
        static int[] channel2 = new int[5];
        static int[] channel3 = new int[5];
        static int[] channel4 = new int[5];

        // No need for array
        static int GPIBAddress = 18;             
        static int TotNumDUTs = 1;                  // Total number of DUTs

        // Need for array
        static int[] channelSize = new int[5];      // Number of WGFMU channels available

        // Need for array
        static int DUTNum1 = 1;                      // DUT number
        static int DUTNum2 = 2;
        static int DUTNum3 = 3;
        static int DUTNum4 = 4;
        static int DUTNum5 = 5;
        static int[] DutNum = new int[5];

        // DC parameters
        // Need for array
        static double vGateStress1 = -1e10;          // Gate Stress Volage - default give as an error
        static double vGateStress2 = -1e10;
        static double vGateStress3 = -1e10;
        static double vGateStress4 = -1e10;
        static double vGateStress5 = -1e10;
        static double[] vGateStress = new double[5];

        static double vGateSense1 = -1e10;           // Gate sense voltage - give as an error
        static double vGateSense2 = -1e10;
        static double vGateSense3 = -1e10;
        static double vGateSense4 = -1e10;
        static double vGateSense5 = -1e10;
        static double[] vGateSense = new double[5];

        static double vDrainStress1 = 0;             // Drain stress voltage
        static double vDrainStress2 = 0;
        static double vDrainStress3 = 0;
        static double vDrainStress4 = 0;
        static double vDrainStress5 = 0;
        static double[] vDrainStress = new double[5];

        static double vDrainSense1 = -0.050;         // Drain sense voltage
        static double vDrainSense2 = -0.050;
        static double vDrainSense3 = -0.050;
        static double vDrainSense4 = -0.050;
        static double vDrainSense5 = -0.050;
        static double[] vDrainSense = new double[5];

        static double measDelay1 = 0;                // How long to wait before measuring
        static double measDelay2 = 0;
        static double measDelay3 = 0;
        static double measDelay4 = 0;
        static double measDelay5 = 0;
        static double[] measDelay = new double[5];

        static int measPoints1 = 100;                // number of points to measure
        static int measPoints2 = 100;
        static int measPoints3 = 100;
        static int measPoints4 = 100;
        static int measPoints5 = 100;
        static int[] measPoints = new int[5];

        static int startAvgPoint1 = 20;              // Start point to begin averaging
        static int startAvgPoint2 = 20;
        static int startAvgPoint3 = 20;
        static int startAvgPoint4 = 20;
        static int startAvgPoint5 = 20;
        static int[] startAvgPoint = new int[5];

        static int stopAvgPoint1 = 80;               // End point for averaging
        static int stopAvgPoint2 = 80;
        static int stopAvgPoint3 = 80;
        static int stopAvgPoint4 = 80;
        static int stopAvgPoint5 = 80;
        static int[] stopAvgPoint = new int[5];

        static double sampleInterval1 = 1e-7;        // Time between measurement points
        static double sampleInterval2 = 1e-7;
        static double sampleInterval3 = 1e-7;
        static double sampleInterval4 = 1e-7;
        static double sampleInterval5 = 1e-7;
        static double[] sampleInterval = new double[5];

        static double stressTime1 = 1000;            // How long to stress
        static double stressTime2 = 1000;
        static double stressTime3 = 1000;
        static double stressTime4 = 1000;
        static double stressTime5 = 1000;
        static double[] stressTime = new double[5];

        static int ppd1 = 3;                         // How many points to measure in a decade (points per decade)
        static int ppd2 = 3;
        static int ppd3 = 3;
        static int ppd4 = 3;
        static int ppd5 = 3;
        static int[] ppd = new int[5];

        static bool isLog1 = true;                   // Measure log or linearly
        static bool isLog2 = true;
        static bool isLog3 = true;
        static bool isLog4 = true;
        static bool isLog5 = true;
        static bool[] isLog = new bool[5];

        static double stepTime1 = 1;                 // How often to sample in linear measurement space
        static double stepTime2 = 1;
        static double stepTime3 = 1;
        static double stepTime4 = 1;
        static double stepTime5 = 1;
        static double[] stepTime = new double[5];

        static double relaxTime1 = 0;                // how long to perform the relaxation part of the test
        static double relaxTime2 = 0;
        static double relaxTime3 = 0;
        static double relaxTime4 = 0;
        static double relaxTime5 = 0;
        static double[] relaxTime = new double[5];

        static double vGateRelax1 = 0;               // the gate voltage during relaxation
        static double vGateRelax2 = 0;
        static double vGateRelax3 = 0;
        static double vGateRelax4 = 0;
        static double vGateRelax5 = 0;
        static double[] vGateRelax = new double[5];

        static double vDrainRelax1 = 0;              // the drain voltage during relaxation
        static double vDrainRelax2 = 0;
        static double vDrainRelax3 = 0;
        static double vDrainRelax4 = 0;
        static double vDrainRelax5 = 0;
        static double[] vDrainRelax = new double[5];

        static double initialSenseTime1 = 10e-6;     // How long to wait before the first measurement point, which must be laregr than or eqalo period1 = 1/freq1 for both AC and DC
        static double initialSenseTime2 = 10e-6;
        static double initialSenseTime3 = 10e-6;
        static double initialSenseTime4 = 10e-6;
        static double initialSenseTime5 = 10e-6;
        static double[] initialSenseTime = new double[5];

        static double gateTransTime1 = 10e-9;        // Rising and falling edge of the gate pulse
        static double gateTransTime2 = 10e-9;
        static double gateTransTime3 = 10e-9;
        static double gateTransTime4 = 10e-9;
        static double gateTransTime5 = 10e-9;
        static double[] gateTransTime = new double[5];

        static double drainTransTime1 = 10e-9;       // Rising and falling edge of the drain pulse
        static double drainTransTime2 = 10e-9;
        static double drainTransTime3 = 10e-9;
        static double drainTransTime4 = 10e-9;
        static double drainTransTime5 = 10e-9;
        static double[] drainTransTime = new double[5];

        static string savePath1 = "";                // where to save the data
        static string savePath2 = "";
        static string savePath3 = "";
        static string savePath4 = "";
        static string savePath5 = "";
        static string[] savePath = new string[5];

        static double measIRange1 = 1e-3;            // the current measurement range to use (on the drain)
        static double measIRange2 = 1e-3;
        static double measIRange3 = 1e-3;
        static double measIRange4 = 1e-3;
        static double measIRange5 = 1e-3;
        static double[] measIRange = new double[5];

        static double maxPulseWidth1 = 3600;         // the max stress/relax time before making a measurement
        static double maxPulseWidth2 = 3600;
        static double maxPulseWidth3 = 3600;
        static double maxPulseWidth4 = 3600;
        static double maxPulseWidth5 = 3600;
        static double[] maxPulseWidth = new double[5];

        // AC parameters
        // Need for array
        static bool acStress1 = false;               // Flag to perform AC stress or not
        static bool acStress2 = false;
        static bool acStress3 = false;
        static bool acStress4 = false;
        static bool acStress5 = false;
        static bool[] acStress = new bool[5];

        static double vGateACLow1 = 0;               // The gate voltage when the AC pulse is on the low cycle (vGateStress1 is the high pulse AC cycle voltage)
        static double vGateACLow2 = 0;
        static double vGateACLow3 = 0;
        static double vGateACLow4 = 0;
        static double vGateACLow5 = 0;
        static double[] vGateACLow = new double[5];

        static double vDrainACLow1 = 0;              // The drain voltage when the AC pulse is on the low cycle (vDrainStress1 is the high pulse AC cycle voltage)
        static double vDrainACLow2 = 0;
        static double vDrainACLow3 = 0;
        static double vDrainACLow4 = 0;
        static double vDrainACLow5 = 0;
        static double[] vDrainACLow = new double[5];

        static bool invStress1 = false;              // Drain low when gate high
        static bool invStress2 = false;
        static bool invStress3 = false;
        static bool invStress4 = false;
        static bool invStress5 = false;
        static bool[] invStress = new bool[5];

        static double skew1 = 0;                     // controls the skew1 of the drain relative to the gate symmetrically skews on both sides of the pulse (0 = overlapping pulses, -10e-9 = drain transisiton 10ns before gate, 10e-9 = drain 10ns after gate)
        static double skew2 = 0;
        static double skew3 = 0;
        static double skew4 = 0;
        static double skew5 = 0;
        static double[] skew = new double[5];

        static double freq1 = 50000;                 // The AC stress frequency
        static double freq2 = 50000;
        static double freq3 = 50000;
        static double freq4 = 50000;
        static double freq5 = 50000;
        static double[] freq = new double[5];

        static double dutyCycle1 = 50;               // Duty cyle of the pulse in percent
        static double dutyCycle2 = 50;
        static double dutyCycle3 = 50;
        static double dutyCycle4 = 50;
        static double dutyCycle5 = 50;
        static double[] dutyCycle = new double[5];

        // these will be recalculated...
        static double period1 = 1 / freq1;
        static double period2 = 1 / freq2;
        static double period3 = 1 / freq3;
        static double period4 = 1 / freq4;
        static double period5 = 1 / freq5;
        static double[] period = new double[5];

        static double gateLowTime1 = (100 - dutyCycle1) * 0.01 * period1 - gateTransTime1;
        static double gateLowTime2 = (100 - dutyCycle2) * 0.01 * period2 - gateTransTime2;
        static double gateLowTime3 = (100 - dutyCycle3) * 0.01 * period3 - gateTransTime3;
        static double gateLowTime4 = (100 - dutyCycle4) * 0.01 * period4 - gateTransTime4;
        static double gateLowTime5 = (100 - dutyCycle5) * 0.01 * period5 - gateTransTime5;
        static double[] gateLowTime = new double[5];

        static double drainHighTime1 = dutyCycle1 * 0.01 * period1 - drainTransTime1;// - 2 * skew1;
        static double drainHighTime2 = dutyCycle2 * 0.01 * period2 - drainTransTime2;
        static double drainHighTime3 = dutyCycle3 * 0.01 * period3 - drainTransTime3;
        static double drainHighTime4 = dutyCycle4 * 0.01 * period4 - drainTransTime4;
        static double drainHighTime5 = dutyCycle5 * 0.01 * period5 - drainTransTime5;
        static double[] drainHighTime = new double[5];

        static double gateHighTime1 = dutyCycle1 * 0.01 * period1 - gateTransTime1;
        static double gateHighTime2 = dutyCycle2 * 0.01 * period2 - gateTransTime2;
        static double gateHighTime3 = dutyCycle3 * 0.01 * period3 - gateTransTime3;
        static double gateHighTime4 = dutyCycle4 * 0.01 * period4 - gateTransTime4;
        static double gateHighTime5 = dutyCycle5 * 0.01 * period5 - gateTransTime5;
        static double[] gateHighTime = new double[5];

        static double drainLowTime1 = (100 - dutyCycle1) * 0.01 * period1 - drainTransTime1;//+ 2 * skew1;
        static double drainLowTime2 = (100 - dutyCycle2) * 0.01 * period2 - drainTransTime2;
        static double drainLowTime3 = (100 - dutyCycle3) * 0.01 * period3 - drainTransTime3;
        static double drainLowTime4 = (100 - dutyCycle4) * 0.01 * period4 - drainTransTime4;
        static double drainLowTime5 = (100 - dutyCycle5) * 0.01 * period5 - drainTransTime5;
        static double[] drainLowTime = new double[5];

        static bool measAfterHigh1 = false;          // measure the current after the high side of the pulse instead of the low - not yet implemented
        static bool measAfterHigh2 = false;
        static bool measAfterHigh3 = false;
        static bool measAfterHigh4 = false;
        static bool measAfterHigh5 = false;
        static bool[] measAfterHigh = new bool[5];

        // Charge pumping parameters - we assume that we cannot measure Idrain so cp will be done during the sense phase - not yet implemented
        // Need for array
        static bool measCP1 = false;                 // Flag to tell th program to perform chargepumping during the sense phase of the test
        static bool measCP2 = false;
        static bool measCP3 = false;
        static bool measCP4 = false;
        static bool measCP5 = false;
        static bool[] measCP = new bool[5];
    
        static double cpFreqStart1 = 1e6;            // The start frequency of the CP freq1 sweep
        static double cpFreqStart2 = 1e6;
        static double cpFreqStart3 = 1e6;
        static double cpFreqStart4 = 1e6;
        static double cpFreqStart5 = 1e6;
        static double[] cpFreqStart = new double[5];

        static double cpFreqStep1 = 4e6;             // The step frequency of the CP freq1 sweep
        static double cpFreqStep2 = 4e6;
        static double cpFreqStep3 = 4e6;
        static double cpFreqStep4 = 4e6;
        static double cpFreqStep5 = 4e6;
        static double[] cpFreqStep = new double[5];

        static int cpNumSteps1 = 1;               // The number of the CP freq1 steps to take
        static int cpNumSteps2 = 1;
        static int cpNumSteps3 = 1;
        static int cpNumSteps4 = 1;
        static int cpNumSteps5 = 1;
        static int[] cpNumSteps = new int[5];

        static double cpHigh1 = 1;                   // The high voltage of the cp
        static double cpHigh2 = 1;
        static double cpHigh3 = 1;
        static double cpHigh4 = 1;
        static double cpHigh5 = 1;
        static double[] cpHigh = new double[5];

        static double cpLow1 = -1;                   // The low voltage of the cp
        static double cpLow2 = -1;
        static double cpLow3 = -1;
        static double cpLow4 = -1;
        static double cpLow5 = -1;
        static double[] cpLow = new double[5];

        static double cpTrans1 = 50e-9;              // The transistion time of cp pulse
        static double cpTrans2 = 50e-9;
        static double cpTrans3 = 50e-9;
        static double cpTrans4 = 50e-9;
        static double cpTrans5 = 50e-9;
        static double[] cpTrans = new double[5];

        // ivMeas parameters (for making and IV during the measurement part of the test)
        // Need for array
        static bool measIV1 = false;
        static bool measIV2 = false;
        static bool measIV3 = false;
        static bool measIV4 = false;
        static bool measIV5 = false;
        static bool[] measIV = new bool[5];

        static double ivGateStart1 = 0;
        static double ivGateStart2 = 0;
        static double ivGateStart3 = 0;
        static double ivGateStart4 = 0;
        static double ivGateStart5 = 0;
        static double[] ivGateStart = new double[5];

        static double ivGateStop1 = 0.5;
        static double ivGateStop2 = 0.5;
        static double ivGateStop3 = 0.5;
        static double ivGateStop4 = 0.5;
        static double ivGateStop5 = 0.5;
        static double[] ivGateStop = new double[5];

        static double ivGateStep1 = 0.05;
        static double ivGateStep2 = 0.05;
        static double ivGateStep3 = 0.05;
        static double ivGateStep4 = 0.05;
        static double ivGateStep5 = 0.05;
        static double[] ivGateStep = new double[5];

        // pMeas parameters (for making arbitrary spot measurements)
        // Need for array
        static bool measSpot1 = false;
        static bool measSpot2 = false;
        static bool measSpot3 = false;
        static bool measSpot4 = false;
        static bool measSpot5 = false;
        static bool[] measSpot = new bool[5];

        static List<double> pMeasGate = new List<double>();
        static List<double> pMeasDrain = new List<double>();

        // here we hold the stress and relaxation times
        static List<double> stressTimePoint;
        static List<double> temp_stressTimePoint1;
        static List<double> temp_stressTimePoint2;
        static List<double> temp_stressTimePoint3;
        static List<double> temp_stressTimePoint4;
        static List<double> temp_stressTimePoint5;
        
        static List<double> relaxTimePoint;
        static List<double> temp_relaxTimePoint1;
        static List<double> temp_relaxTimePoint2;
        static List<double> temp_relaxTimePoint3;
        static List<double> temp_relaxTimePoint4;
        static List<double> temp_relaxTimePoint5;

        static int measWindowsPerSense1 = 1; // holds how many measurement windows per sense (e.g., CP three freq1 would have 3 meas windows. or an IV sweep could have many more)
        static int measWindowsPerSense2 = 1;
        static int measWindowsPerSense3 = 1;
        static int measWindowsPerSense4 = 1;
        static int measWindowsPerSense5 = 1;
        static int[] measWindowsPerSense = new int[5];

        // Cycle stress - repeat the stress relaxation a number of cycles - Not yet implemented
        static int numCycles = 1;                   // Default - perform 1 cycle

        static void Main(string[] args)
        {
            // Parse the arguments
            if (parseArguments(args) == true)
            {
                // Must be defined after parseArguments
                for (int i = 0; i < TotNumDUTs; i++)
                {
                    if (i == 0)
                    {
                        maxPulseWidth[i] = maxPulseWidth1;
                        measWindowsPerSense[i] = measWindowsPerSense1;

                        if (channel1[i] > 500)
                            channelSize[i]++;
                        if (channel2[i] > 500)
                            channelSize[i]++;
                        if (channel3[i] > 500)
                            channelSize[i]++;
                        if (channel4[i] > 500)
                            channelSize[i]++;
                    }
                        
                    else if (i == 1)
                    {
                        maxPulseWidth[i] = maxPulseWidth2;
                        measWindowsPerSense[i] = measWindowsPerSense2;

                        if (channel1[i] > 500)
                            channelSize[i]++;
                        if (channel2[i] > 500)
                            channelSize[i]++;
                        if (channel3[i] > 500)
                            channelSize[i]++;
                        if (channel4[i] > 500)
                            channelSize[i]++;
                    }
                        
                    else if (i == 2)
                    {
                        maxPulseWidth[i] = maxPulseWidth3;
                        measWindowsPerSense[i] = measWindowsPerSense3;

                        if (channel1[i] > 500)
                            channelSize[i]++;
                        if (channel2[i] > 500)
                            channelSize[i]++;
                        if (channel3[i] > 500)
                            channelSize[i]++;
                        if (channel4[i] > 500)
                            channelSize[i]++;
                    }
                        
                    else if (i == 3)
                    {
                        maxPulseWidth[i] = maxPulseWidth4;
                        measWindowsPerSense[i] = measWindowsPerSense4;

                        if (channel1[i] > 500)
                            channelSize[i]++;
                        if (channel2[i] > 500)
                            channelSize[i]++;
                        if (channel3[i] > 500)
                            channelSize[i]++;
                        if (channel4[i] > 500)
                            channelSize[i]++;
                    }
                        
                    else if (i == 4)
                    {
                        maxPulseWidth[i] = maxPulseWidth5;
                        measWindowsPerSense[i] = measWindowsPerSense5;

                        if (channel1[i] > 500)
                            channelSize[i]++;
                        if (channel2[i] > 500)
                            channelSize[i]++;
                        if (channel3[i] > 500)
                            channelSize[i]++;
                        if (channel4[i] > 500)
                            channelSize[i]++;
                    }                        
                }

                Console.WriteLine("...............DUT Under Stress....................");
                Console.WriteLine("Total # of DUTs under stress = " + TotNumDUTs);
                Console.WriteLine("");

                for (int i = 0; i < TotNumDUTs; i++)
                    writeParamsToConsole(i);

                Console.WriteLine("..................................................................");
                Console.WriteLine("........... Input parameters are succefully processed ............");
                Console.WriteLine("..................................................................");

                WGFMU.openLogFile("C:\\WGFMU.log");
                WGFMU.clear();
                WGFMU.openSession("GPIB0::" + GPIBAddress.ToString() + "::INSTR"); // ANT uses 18 as the address here we enable to option to change to whatever
                WGFMU.initialize();

                // Now lets findout how many channels we have to work with
                // User will set the channelSize from Visual Basic
                // WGFMU.getChannelIdSize(ref channelSize);

                for (int i = 0; i < TotNumDUTs; i++)
                    takeData(i);

                // Connect all available channels
                for (int i = 0; i < TotNumDUTs; i++)
                {
                    WGFMU.connect(channel1[i]);
                    WGFMU.connect(channel2[i]);

                    if (channelSize[i] > 2)
                    {
                        WGFMU.connect(channel3[i]);

                        if (channelSize[i] > 3)
                            WGFMU.connect(channel4[i]);
                    }
                }

                //--------------------------------------------------------------------
                // Find longest stress and relax time for timing purpose.
                int index = 0;
                double time = stressTime[0] + relaxTime[0];

                for (int i = 1; i < TotNumDUTs; i++)
                {
                    if (time < (stressTime[i] + relaxTime[i]))
                    {
                        time = stressTime[i] + relaxTime[i];
                        index = i;
                    }                        
                }
                
                Console.WriteLine("");
                Console.WriteLine("*** Total number of DUTs in parallel stress = " + TotNumDUTs.ToString() + " ***");
                Console.WriteLine("*** Displaying measurement status of only DUT#" + DutNum[index].ToString() + " ***");
                Console.WriteLine("*** (longest total time = stress time + relax time) ***");
                WGFMU.execute();
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = 0.0/{2:F1}", 0, stressTime[index], relaxTime[index]);

                startTime = DateTime.Now;

                // Create a timer with a two second interval.
                Timer aTimer = new System.Timers.Timer(2000);
                // Hook up the Elapsed event for the timer.
                if (index == 0)
                    aTimer.Elapsed += OnTimedEvent0;
                else if (index == 1)
                    aTimer.Elapsed += OnTimedEvent1;
                else if (index == 2)
                    aTimer.Elapsed += OnTimedEvent2;
                else if (index == 3)
                    aTimer.Elapsed += OnTimedEvent3;
                else if (index == 4)
                    aTimer.Elapsed += OnTimedEvent4;

                aTimer.AutoReset = true;
                aTimer.Enabled = true;

                // Had an error during "WGFMU.waitUntilCompleted" not sure why. Lets try passing the time here
                // Lets kick of a thread that will display the time

                int ret = 0;
                bool errorOccured = false;
                ret = WGFMU.waitUntilCompleted();   // there have been times when this fails and we can keep it going...
                // I have also seen that when I keep it going then it doesn't catch
                // that the test has ended

                if (ret < 0)
                {
                    while (aTimer.Enabled == true)
                    {
                        System.Threading.Thread.Sleep(2000);
                        errorOccured = true;
                    }
                }

                endTime = DateTime.Now;

                aTimer.Stop();
                aTimer.Dispose();

                //--------------------------------------------------------------------
                // Disconnect all available channels
                for (int i = 0; i < TotNumDUTs; i++)
                {
                    // these functions were added (perhaps WGFMU.initialize() does the same thing)
                    WGFMU.disconnect(channel1[i]);
                    WGFMU.disconnect(channel2[i]);

                    if (channelSize[i] > 2)
                    {
                        WGFMU.disconnect(channel3[i]);

                        if (channelSize[i] > 3)
                            WGFMU.disconnect(channel4[i]);
                    }
                }

                // Write all measurements from WGFMU to files

                for (int i = 0; i < TotNumDUTs; i++)
                {
                    Console.WriteLine("");
                    Console.WriteLine("Writing results for DUT#" + DutNum[i].ToString() + "..........");

                    if(i == 0)
                        writeResults(i, channel2[i], savePath[i], channel1[i], temp_stressTimePoint1, temp_relaxTimePoint1);
                    else if(i == 1)
                        writeResults(i, channel2[i], savePath[i], channel1[i], temp_stressTimePoint2, temp_relaxTimePoint2);
                    else if(i == 2)
                        writeResults(i, channel2[i], savePath[i], channel1[i], temp_stressTimePoint3, temp_relaxTimePoint3);
                    else if(i == 3)
                        writeResults(i, channel2[i], savePath[i], channel1[i], temp_stressTimePoint4, temp_relaxTimePoint4);
                    else if(i == 4)
                        writeResults(i, channel2[i], savePath[i], channel1[i], temp_stressTimePoint5, temp_relaxTimePoint5);
                }

                WGFMU.initialize();
                WGFMU.closeSession();
                WGFMU.closeLogFile();

                Console.WriteLine("");
                Console.WriteLine("..........Finished all DUTs measurements..........");
            }

            else
            {
                //Print error in arguments
                {
                    string argLine = "";
                    for (int i = 0; i < args.Length; i++)
                    {
                        argLine += args[i];
                    }

                    string[] lines = { "Error taking user input", "Passed Arguments are", argLine };

                    for(int i = 0; i < TotNumDUTs; i++)
                    {
                        System.IO.File.WriteAllLines(@savePath[i], lines);
                    }
                }
            }
        }

        static bool parseArguments(string[] args)
        {
            if (args.Length < 2) return false;

            string str = "";

            // build the string back
            for (int i = 0; i < args.Length; i++)
            {
                str += args[i];
            }

            args = str.Split('~');

            string tempParam = "";
            try
            {
                for (int i = 0; i < args.Length; i++)
                {
                    tempParam = args[i];
                    string[] parsedStr = args[i].Split('=');
                    string[] parsedStrColon = args[i].Split(':');

                    if (parsedStrColon[0] == "-save1")
                    {
                        savePath1 = parsedStrColon[1] + ":" + parsedStrColon[2];
                        savePath[0] = savePath1;
                    }

                    else if (parsedStrColon[0] == "-save2")
                    {
                        savePath2 = parsedStrColon[1] + ":" + parsedStrColon[2];
                        savePath[1] = savePath2;
                    }

                    else if (parsedStrColon[0] == "-save3")
                    {
                        savePath3 = parsedStrColon[1] + ":" + parsedStrColon[2];
                        savePath[2] = savePath3;
                    }

                    else if (parsedStrColon[0] == "-save4")
                    {
                        savePath4 = parsedStrColon[1] + ":" + parsedStrColon[2];
                        savePath[3] = savePath4;
                    }

                    else if (parsedStrColon[0] == "-save5")
                    {
                        savePath5 = parsedStrColon[1] + ":" + parsedStrColon[2];
                        savePath[4] = savePath5;
                    }

                    else
                    {
                        switch (parsedStr[0])
                        {
                            case "TotNumDUTs":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out TotNumDUTs);
                                break;
                            //----------------------------------------------------                           
                            case "DUT1_channel1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel1[0]);
                                break;
                            case "DUT1_channel2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel2[0]);
                                break;
                            case "DUT1_channel3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel3[0]);
                                break;
                            case "DUT1_channel4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel4[0]);
                                break;
                            case "DUT2_channel1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel1[1]);
                                break;
                            case "DUT2_channel2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel2[1]);
                                break;
                            case "DUT2_channel3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel3[1]);
                                break;
                            case "DUT2_channel4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel4[1]);
                                break;
                            case "DUT3_channel1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel1[2]);
                                break;
                            case "DUT3_channel2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel2[2]);
                                break;
                            case "DUT3_channel3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel3[2]);
                                break;
                            case "DUT3_channel4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel4[2]);
                                break;
                            case "DUT4_channel1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel1[3]);
                                break;
                            case "DUT4_channel2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel2[3]);
                                break;
                            case "DUT4_channel3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel3[3]);
                                break;
                            case "DUT4_channel4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel4[3]);
                                break;
                            case "DUT5_channel1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel1[4]);
                                break;
                            case "DUT5_channel2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel2[4]);
                                break;
                            case "DUT5_channel3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel3[4]);
                                break;
                            case "DUT5_channel4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out channel4[4]);
                                break;
                            //----------------------------------------------------
                            case "DUTNum1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out DUTNum1);
                                    DutNum[0] = DUTNum1;
                                } 
                                break;
                            case "DUTNum2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out DUTNum2);
                                    DutNum[1] = DUTNum2;
                                }
                                break;
                            case "DUTNum3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out DUTNum3);
                                    DutNum[2] = DUTNum3;
                                }
                                break;
                            case "DUTNum4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out DUTNum4);
                                    DutNum[3] = DUTNum4;
                                }
                                break;
                            case "DUTNum5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out DUTNum5);
                                    DutNum[4] = DUTNum5;
                                }
                                break;
                            //----------------------------------------------------
                            case "GPIBAddress":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out GPIBAddress);
                                break;
                            //----------------------------------------------------
                            case "VGateStress1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateStress1);
                                    vGateStress[0] = vGateStress1;
                                }
                                break;
                            case "VGateStress2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateStress2);
                                    vGateStress[1] = vGateStress2;
                                }   
                                break;
                            case "VGateStress3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateStress3);
                                    vGateStress[2] = vGateStress3;
                                }   
                                break;
                            case "VGateStress4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateStress4);
                                    vGateStress[3] = vGateStress4;
                                }
                                break;
                            case "VGateStress5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateStress5);
                                    vGateStress[4] = vGateStress5;
                                }
                                break;
                            //----------------------------------------------------
                            case "VGateSense1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateSense1);
                                    vGateSense[0] = vGateSense1;
                                } 
                                break;
                            case "VGateSense2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateSense2);
                                    vGateSense[1] = vGateSense2;
                                }
                                break;
                            case "VGateSense3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateSense3);
                                    vGateSense[2] = vGateSense3;
                                }
                                break;
                            case "VGateSense4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateSense4);
                                    vGateSense[3] = vGateSense4;
                                }
                                break;
                            case "VGateSense5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateSense5);
                                    vGateSense[4] = vGateSense5;
                                }
                                break;
                            //----------------------------------------------------
                            case "VDrainStress1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainStress1);
                                    vDrainStress[0] = vDrainStress1;
                                }  
                                break;
                            case "VDrainStress2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainStress2);
                                    vDrainStress[1] = vDrainStress2;
                                }
                                break;
                            case "VDrainStress3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainStress3);
                                    vDrainStress[2] = vDrainStress3;
                                }
                                break;
                            case "VDrainStress4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainStress4);
                                    vDrainStress[3] = vDrainStress4;
                                }
                                break;
                            case "VDrainStress5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainStress5);
                                    vDrainStress[4] = vDrainStress5;
                                }
                                break;
                            //----------------------------------------------------
                            case "VDrainSense1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainSense1);
                                    vDrainSense[0] = vDrainSense1;
                                }
                                break;
                            case "VDrainSense2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainSense2);
                                    vDrainSense[1] = vDrainSense2;
                                }
                                break;
                            case "VDrainSense3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainSense3);
                                    vDrainSense[2] = vDrainSense3;
                                }
                                break;
                            case "VDrainSense4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainSense4);
                                    vDrainSense[3] = vDrainSense4;
                                }
                                break;
                            case "VDrainSense5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainSense5);
                                    vDrainSense[4] = vDrainSense5;
                                }
                                break;
                            //----------------------------------------------------
                            case "measDelay1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measDelay1);
                                    measDelay[0] = measDelay1;
                                }
                                break;
                            case "measDelay2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measDelay2);
                                    measDelay[1] = measDelay2;
                                }
                                break;
                            case "measDelay3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measDelay3);
                                    measDelay[2] = measDelay3;
                                }
                                break;
                            case "measDelay4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measDelay4);
                                    measDelay[3] = measDelay4;
                                }
                                break;
                            case "measDelay5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measDelay5);
                                    measDelay[4] = measDelay5;
                                }
                                break;
                            //----------------------------------------------------
                            case "measPoints1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out measPoints1);
                                    stopAvgPoint1 = (int)(0.8 * measPoints1);     // We will always default to 20 to 80% (need to spec start and stop after spec measPoints1)
                                    startAvgPoint1 = (int)(0.2 * measPoints1);
                                    measPoints[0] = measPoints1;
                                }
                                break;
                            case "measPoints2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out measPoints2);
                                    stopAvgPoint2 = (int)(0.8 * measPoints2);     // We will always default to 20 to 80% (need to spec start and stop after spec measPoints1)
                                    startAvgPoint2 = (int)(0.2 * measPoints2);
                                    measPoints[1] = measPoints2;
                                }
                                break;
                            case "measPoints3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out measPoints3);
                                    stopAvgPoint3 = (int)(0.8 * measPoints3);     // We will always default to 20 to 80% (need to spec start and stop after spec measPoints1)
                                    startAvgPoint3 = (int)(0.2 * measPoints3);
                                    measPoints[2] = measPoints3;
                                }
                                break;
                            case "measPoints4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out measPoints4);
                                    stopAvgPoint4 = (int)(0.8 * measPoints4);     // We will always default to 20 to 80% (need to spec start and stop after spec measPoints1)
                                    startAvgPoint4 = (int)(0.2 * measPoints4);
                                    measPoints[3] = measPoints4;
                                }
                                break;
                            case "measPoints5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out measPoints5);
                                    stopAvgPoint5 = (int)(0.8 * measPoints5);     // We will always default to 20 to 80% (need to spec start and stop after spec measPoints1)
                                    startAvgPoint5 = (int)(0.2 * measPoints5);
                                    measPoints[4] = measPoints5;
                                }
                                break;
                            //----------------------------------------------------
                            case "startAvgPoint1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out startAvgPoint1);
                                    if (startAvgPoint1 < 1) startAvgPoint1 = 1;
                                    if (startAvgPoint1 > measPoints1) startAvgPoint1 = measPoints1;
                                    startAvgPoint[0] = startAvgPoint1;
                                }
                                break;
                            case "startAvgPoint2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out startAvgPoint2);
                                    if (startAvgPoint2 < 1) startAvgPoint2 = 1;
                                    if (startAvgPoint2 > measPoints2) startAvgPoint2 = measPoints2;
                                    startAvgPoint[1] = startAvgPoint2;
                                }
                                break;
                            case "startAvgPoint3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out startAvgPoint3);
                                    if (startAvgPoint3 < 1) startAvgPoint3 = 1;
                                    if (startAvgPoint3 > measPoints3) startAvgPoint3 = measPoints3;
                                    startAvgPoint[2] = startAvgPoint3;
                                }
                                break;
                            case "startAvgPoint4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out startAvgPoint4);
                                    if (startAvgPoint4 < 1) startAvgPoint4 = 1;
                                    if (startAvgPoint4 > measPoints4) startAvgPoint4 = measPoints4;
                                    startAvgPoint[3] = startAvgPoint4;
                                }
                                break;
                            case "startAvgPoint5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out startAvgPoint5);
                                    if (startAvgPoint5 < 1) startAvgPoint5 = 1;
                                    if (startAvgPoint5 > measPoints5) startAvgPoint5 = measPoints5;
                                    startAvgPoint[4] = startAvgPoint5;
                                }
                                break;
                            //----------------------------------------------------
                            case "stopAvgPoint1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out stopAvgPoint1);
                                    if (stopAvgPoint1 > measPoints1) stopAvgPoint1 = measPoints1;
                                    if (stopAvgPoint1 < startAvgPoint1) stopAvgPoint1 = startAvgPoint1;
                                    stopAvgPoint[0] = stopAvgPoint1;
                                }
                                break;
                            case "stopAvgPoint2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out stopAvgPoint2);
                                    if (stopAvgPoint2 > measPoints2) stopAvgPoint2 = measPoints2;
                                    if (stopAvgPoint2 < startAvgPoint2) stopAvgPoint2 = startAvgPoint2;
                                    stopAvgPoint[1] = stopAvgPoint2;
                                }
                                break;
                            case "stopAvgPoint3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out stopAvgPoint3);
                                    if (stopAvgPoint3 > measPoints3) stopAvgPoint3 = measPoints3;
                                    if (stopAvgPoint3 < startAvgPoint3) stopAvgPoint3 = startAvgPoint3;
                                    stopAvgPoint[2] = stopAvgPoint3;
                                }
                                break;
                            case "stopAvgPoint4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out stopAvgPoint4);
                                    if (stopAvgPoint4 > measPoints4) stopAvgPoint4 = measPoints4;
                                    if (stopAvgPoint4 < startAvgPoint4) stopAvgPoint4 = startAvgPoint4;
                                    stopAvgPoint[3] = stopAvgPoint4;
                                }
                                break;
                            case "stopAvgPoint5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out stopAvgPoint5);
                                    if (stopAvgPoint5 > measPoints5) stopAvgPoint5 = measPoints5;
                                    if (stopAvgPoint5 < startAvgPoint5) stopAvgPoint5 = startAvgPoint5;
                                    stopAvgPoint[4] = stopAvgPoint5;
                                }
                                break;
                            //----------------------------------------------------
                            case "sampleInterval1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out sampleInterval1);
                                    sampleInterval[0] = sampleInterval1;
                                }
                                break;
                            case "sampleInterval2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out sampleInterval2);
                                    sampleInterval[1] = sampleInterval2;
                                }
                                break;
                            case "sampleInterval3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out sampleInterval3);
                                    sampleInterval[2] = sampleInterval3;
                                }
                                break;
                            case "sampleInterval4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out sampleInterval4);
                                    sampleInterval[3] = sampleInterval4;
                                }
                                break;
                            case "sampleInterval5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out sampleInterval5);
                                    sampleInterval[4] = sampleInterval5;
                                }
                                break;
                            //----------------------------------------------------
                            case "stressTime1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stressTime1);
                                    stressTime[0] = stressTime1;
                                }
                                break;
                            case "stressTime2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stressTime2);
                                    stressTime[1] = stressTime2;
                                }
                                break;
                            case "stressTime3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stressTime3);
                                    stressTime[2] = stressTime3;
                                }
                                break;
                            case "stressTime4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stressTime4);
                                    stressTime[3] = stressTime4;
                                }
                                break;
                            case "stressTime5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stressTime5);
                                    stressTime[4] = stressTime5;
                                }
                                break;
                            //----------------------------------------------------
                            case "ppd1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out ppd1);
                                    ppd[0] = ppd1;
                                }
                                break;
                            case "ppd2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out ppd2);
                                    ppd[1] = ppd2;
                                }
                                break;
                            case "ppd3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out ppd3);
                                    ppd[2] = ppd3;
                                }
                                break;
                            case "ppd4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out ppd4);
                                    ppd[3] = ppd4;
                                }
                                break;
                            case "ppd5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out ppd5);
                                    ppd[4] = ppd5;
                                }
                                break;
                            //----------------------------------------------------
                            case "stepTime1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stepTime1);
                                    stepTime[0] = stepTime1;
                                }
                                break;
                            case "stepTime2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stepTime2);
                                    stepTime[1] = stepTime2;
                                }
                                break;
                            case "stepTime3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stepTime3);
                                    stepTime[2] = stepTime3;
                                }
                                break;
                            case "stepTime4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stepTime4);
                                    stepTime[3] = stepTime4;
                                }
                                break;
                            case "stepTime5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out stepTime5);
                                    stepTime[4] = stepTime5;
                                }
                                break;
                            //----------------------------------------------------
                            case "isLog1":
                                if (parsedStr[1].Length == 0)
                                {
                                    isLog1 = false;
                                    isLog[0] = isLog1;
                                    break;
                                }
                                else
                                {
                                    isLog1 = Convert.ToBoolean(parsedStr[1]);
                                    isLog[0] = isLog1;
                                }
                                break;
                            case "isLog2":
                                if (parsedStr[1].Length == 0)
                                {
                                    isLog2 = false;
                                    isLog[1] = isLog2;
                                    break;
                                }
                                else
                                {
                                    isLog2 = Convert.ToBoolean(parsedStr[1]);
                                    isLog[1] = isLog2;
                                }
                                break;
                            case "isLog3":
                                if (parsedStr[1].Length == 0)
                                {
                                    isLog3 = false;
                                    isLog[2] = isLog3;
                                    break;
                                }
                                else
                                {
                                    isLog3 = Convert.ToBoolean(parsedStr[1]);
                                    isLog[2] = isLog3;
                                }
                                break;
                            case "isLog4":
                                if (parsedStr[1].Length == 0)
                                {
                                    isLog4 = false;
                                    isLog[3] = isLog4;
                                    break;
                                }
                                else
                                {
                                    isLog4 = Convert.ToBoolean(parsedStr[1]);
                                    isLog[3] = isLog4;
                                }
                                break;
                            case "isLog5":
                                if (parsedStr[1].Length == 0)
                                {
                                    isLog5 = false;
                                    isLog[4] = isLog5;
                                    break;
                                }
                                else
                                {
                                    isLog5 = Convert.ToBoolean(parsedStr[1]);
                                    isLog[4] = isLog5;
                                }
                                break;
                            //---------------------------------------------------------------------------
                            case "maxPulseWidth1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out maxPulseWidth1);
                                    maxPulseWidth[0] = maxPulseWidth1;
                                }
                                break;
                            case "maxPulseWidth2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out maxPulseWidth2);
                                    maxPulseWidth[1] = maxPulseWidth2;
                                }
                                break;
                            case "maxPulseWidth3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out maxPulseWidth3);
                                    maxPulseWidth[2] = maxPulseWidth3;
                                }
                                break;
                            case "maxPulseWidth4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out maxPulseWidth4);
                                    maxPulseWidth[3] = maxPulseWidth4;
                                }
                                break;
                            case "maxPulseWidth5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out maxPulseWidth5);
                                    maxPulseWidth[4] = maxPulseWidth5;
                                }
                                break;
                            //----------------------------------------------------
                            case "VGateRelax1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateRelax1);
                                    vGateRelax[0] = vGateRelax1;
                                }   
                                break;
                            case "VGateRelax2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateRelax2);
                                    vGateRelax[1] = vGateRelax2;
                                }
                                break;
                            case "VGateRelax3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateRelax3);
                                    vGateRelax[2] = vGateRelax3;
                                }
                                break;
                            case "VGateRelax4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateRelax4);
                                    vGateRelax[3] = vGateRelax4;
                                }
                                break;
                            case "VGateRelax5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateRelax5);
                                    vGateRelax[4] = vGateRelax5;
                                }
                                break;
                            //----------------------------------------------------
                            case "VDrainRelax1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainRelax1);
                                    vDrainRelax[0] = vDrainRelax1;
                                }
                                break;
                            case "VDrainRelax2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainRelax2);
                                    vDrainRelax[1] = vDrainRelax2;
                                }
                                break;
                            case "VDrainRelax3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainRelax3);
                                    vDrainRelax[2] = vDrainRelax3;
                                }
                                break;
                            case "VDrainRelax4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainRelax4);
                                    vDrainRelax[3] = vDrainRelax4;
                                }
                                break;
                            case "VDrainRelax5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainRelax5);
                                    vDrainRelax[4] = vDrainRelax5;
                                }
                                break;
                            //----------------------------------------------------
                            case "relaxTime1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out relaxTime1);
                                    relaxTime[0] = relaxTime1;
                                }
                                break;
                            case "relaxTime2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out relaxTime2);
                                    relaxTime[1] = relaxTime2;
                                }
                                break;
                            case "relaxTime3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out relaxTime3);
                                    relaxTime[2] = relaxTime3;
                                }
                                break;
                            case "relaxTime4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out relaxTime4);
                                    relaxTime[3] = relaxTime4;
                                }
                                break;
                            case "relaxTime5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out relaxTime5);
                                    relaxTime[4] = relaxTime5;
                                }
                                break;
                            //----------------------------------------------------
                            case "transTime":
                                double.TryParse(parsedStr[1], out gateTransTime1);
                                double.TryParse(parsedStr[1], out drainTransTime1);
                                break;
                            //----------------------------------------------------
                            case "gateTransTime1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out gateTransTime1);
                                    gateTransTime[0] = gateTransTime1;
                                }
                                break;
                            case "gateTransTime2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out gateTransTime2);
                                    gateTransTime[1] = gateTransTime2;
                                }
                                break;
                            case "gateTransTime3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out gateTransTime3);
                                    gateTransTime[2] = gateTransTime3;
                                }
                                break;
                            case "gateTransTime4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out gateTransTime4);
                                    gateTransTime[3] = gateTransTime4;
                                }
                                break;
                            case "gateTransTime5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out gateTransTime5);
                                    gateTransTime[4] = gateTransTime5;
                                }
                                break;
                            //----------------------------------------------------
                            case "drainTransTime1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out drainTransTime1);
                                    drainTransTime[0] = drainTransTime1;
                                }
                                break;
                            case "drainTransTime2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out drainTransTime2);
                                    drainTransTime[1] = drainTransTime2;
                                }
                                break;
                            case "drainTransTime3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out drainTransTime3);
                                    drainTransTime[2] = drainTransTime3;
                                }
                                break;
                            case "drainTransTime4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out drainTransTime4);
                                    drainTransTime[3] = drainTransTime4;
                                }
                                break;
                            case "drainTransTime5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out drainTransTime5);
                                    drainTransTime[4] = drainTransTime5;
                                }
                                break;
                            //----------------------------------------------------
                            case "initialSenseTime1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out initialSenseTime1);
                                    initialSenseTime[0] = initialSenseTime1;
                                }
                                break;
                            case "initialSenseTime2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out initialSenseTime2);
                                    initialSenseTime[1] = initialSenseTime2;
                                }
                                break;
                            case "initialSenseTime3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out initialSenseTime3);
                                    initialSenseTime[2] = initialSenseTime3;
                                }
                                break;
                            case "initialSenseTime4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out initialSenseTime4);
                                    initialSenseTime[3] = initialSenseTime4;
                                }
                                break;
                            case "initialSenseTime5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out initialSenseTime5);
                                    initialSenseTime[4] = initialSenseTime5;
                                }
                                break;
                            //----------------------------------------------------
                            case "measIRange1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measIRange1);
                                    measIRange[0] = measIRange1;
                                }  
                                break;
                            case "measIRange2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measIRange2);
                                    measIRange[1] = measIRange2;
                                }
                                break;
                            case "measIRange3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measIRange3);
                                    measIRange[2] = measIRange3;
                                }
                                break;
                            case "measIRange4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measIRange4);
                                    measIRange[3] = measIRange4;
                                }
                                break;
                            case "measIRange5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out measIRange5);
                                    measIRange[4] = measIRange5;
                                }
                                break;
                            //----------------------------------------------------
                            case "acStress1":
                                if (parsedStr[1].Length == 0)
                                {
                                    acStress1 = false;
                                    acStress[0] = acStress1;
                                    break;
                                }
                                else
                                {
                                    acStress1 = Convert.ToBoolean(parsedStr[1]);
                                    if (acStress1) acStress1 = true;
                                    if (!acStress1) acStress1 = false;
                                    acStress[0] = acStress1;
                                }
                                break;
                            case "acStress2":
                                if (parsedStr[1].Length == 0)
                                {
                                    acStress2 = false;
                                    acStress[1] = acStress2;
                                    break;
                                }
                                else
                                {
                                    acStress2 = Convert.ToBoolean(parsedStr[1]);
                                    if (acStress2) acStress2 = true;
                                    if (!acStress2) acStress2 = false;
                                    acStress[1] = acStress2;
                                }
                                break;
                            case "acStress3":
                                if (parsedStr[1].Length == 0)
                                {
                                    acStress3 = false;
                                    acStress[2] = acStress3;
                                    break;
                                }
                                else
                                {
                                    acStress3 = Convert.ToBoolean(parsedStr[1]);
                                    if (acStress3) acStress3 = true;
                                    if (!acStress3) acStress3 = false;
                                    acStress[2] = acStress3;
                                }
                                break;
                            case "acStress4":
                                if (parsedStr[1].Length == 0)
                                {
                                    acStress4 = false;
                                    acStress[3] = acStress4;
                                    break;
                                }
                                else
                                {
                                    acStress4 = Convert.ToBoolean(parsedStr[1]);
                                    if (acStress4) acStress4 = true;
                                    if (!acStress4) acStress4 = false;
                                    acStress[3] = acStress4;
                                }
                                break;
                            case "acStress5":
                                if (parsedStr[1].Length == 0)
                                {
                                    acStress5 = false;
                                    acStress[4] = acStress5;
                                    break;
                                }
                                else
                                {
                                    acStress5 = Convert.ToBoolean(parsedStr[1]);
                                    if (acStress5) acStress5 = true;
                                    if (!acStress5) acStress5 = false;
                                    acStress[4] = acStress5;
                                }
                                break;
                            //----------------------------------------------------
                            case "dcStress":
                                acStress1 = false;
                                break;
                            //----------------------------------------------------
                            case "VGateACLow1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateACLow1);
                                    vGateACLow[0] = vGateACLow1;
                                }
                                break;
                            case "VGateACLow2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateACLow2);
                                    vGateACLow[1] = vGateACLow2;
                                }
                                break;
                            case "VGateACLow3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateACLow3);
                                    vGateACLow[2] = vGateACLow3;
                                }
                                break;
                            case "VGateACLow4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateACLow4);
                                    vGateACLow[3] = vGateACLow4;
                                }
                                break;
                            case "VGateACLow5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vGateACLow5);
                                    vGateACLow[4] = vGateACLow5;
                                }
                                break;
                            //----------------------------------------------------
                            case "VDrainACLow1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainACLow1);
                                    vDrainACLow[0] = vDrainACLow1;
                                }
                                break;
                            case "VDrainACLow2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainACLow2);
                                    vDrainACLow[1] = vDrainACLow2;
                                }
                                break;
                            case "VDrainACLow3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainACLow3);
                                    vDrainACLow[2] = vDrainACLow3;
                                }
                                break;
                            case "VDrainACLow4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainACLow4);
                                    vDrainACLow[3] = vDrainACLow4;
                                }
                                break;
                            case "VDrainACLow5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out vDrainACLow5);
                                    vDrainACLow[4] = vDrainACLow5;
                                }
                                break;
                            //----------------------------------------------------
                            case "invStress1":
                                if (parsedStr[1].Length == 0)
                                {
                                    invStress1 = false;
                                    invStress[0] = invStress1;
                                    break;
                                }
                                else
                                {
                                    invStress1 = Convert.ToBoolean(parsedStr[1]);
                                    invStress[0] = invStress1;
                                }                                   
                                break;
                            case "invStress2":
                                if (parsedStr[1].Length == 0)
                                {
                                    invStress2 = false;
                                    invStress[1] = invStress2;
                                    break;
                                }
                                else
                                {
                                    invStress2 = Convert.ToBoolean(parsedStr[1]);
                                    invStress[1] = invStress2;
                                }
                                break;
                            case "invStress3":
                                if (parsedStr[1].Length == 0)
                                {
                                    invStress3 = false;
                                    invStress[2] = invStress3;
                                    break;
                                }
                                else
                                {
                                    invStress3 = Convert.ToBoolean(parsedStr[1]);
                                    invStress[2] = invStress3;
                                }
                                break;
                            case "invStress4":
                                if (parsedStr[1].Length == 0)
                                {
                                    invStress4 = false;
                                    invStress[3] = invStress4;
                                    break;
                                }
                                else
                                {
                                    invStress4 = Convert.ToBoolean(parsedStr[1]);
                                    invStress[3] = invStress4;
                                }
                                break;
                            case "invStress5":
                                if (parsedStr[1].Length == 0)
                                {
                                    invStress5 = false;
                                    invStress[4] = invStress5;
                                    break;
                                }
                                else
                                {
                                    invStress5 = Convert.ToBoolean(parsedStr[1]);
                                    invStress[4] = invStress5;
                                }
                                break;
                            //----------------------------------------------------
                            case "skew1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out skew1);
                                    skew1 = Math.Round(skew1, 8);
                                    skew[0] = skew1;
                                }
                                break;
                            case "skew2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out skew2);
                                    skew2 = Math.Round(skew2, 8);
                                    skew[1] = skew2;
                                }
                                break;
                            case "skew3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out skew3);
                                    skew3 = Math.Round(skew3, 8);
                                    skew[2] = skew3;
                                }
                                break;
                            case "skew4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out skew4);
                                    skew4 = Math.Round(skew4, 8);
                                    skew[3] = skew4;
                                }
                                break;
                            case "skew5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out skew5);
                                    skew5 = Math.Round(skew5, 8);
                                    skew[4] = skew5;
                                }
                                break;
                            //----------------------------------------------------
                            case "freq1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out freq1);
                                    freq[0] = freq1;
                                }                               
                                break;
                            case "freq2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out freq2);
                                    freq[1] = freq2;
                                }
                                break;
                            case "freq3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out freq3);
                                    freq[2] = freq3;
                                }
                                break;
                            case "freq4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out freq4);
                                    freq[3] = freq4;
                                }
                                break;
                            case "freq5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out freq5);
                                    freq[4] = freq5;
                                }
                                break;
                            //----------------------------------------------------
                            case "dutyCycle1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out dutyCycle1);
                                    dutyCycle[0] = dutyCycle1;
                                }                                    
                                break;
                            case "dutyCycle2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out dutyCycle2);
                                    dutyCycle[1] = dutyCycle2;
                                }
                                break;
                            case "dutyCycle3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out dutyCycle3);
                                    dutyCycle[2] = dutyCycle3;
                                }
                                break;
                            case "dutyCycle4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out dutyCycle4);
                                    dutyCycle[3] = dutyCycle4;
                                }
                                break;
                            case "dutyCycle5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out dutyCycle5);
                                    dutyCycle[4] = dutyCycle5;
                                }
                                break;
                            //----------------------------------------------------
                            case "measAfterHigh1":
                                if (parsedStr[1].Length == 0)
                                {
                                    measAfterHigh1 = false;
                                    measAfterHigh[0] = measAfterHigh1;
                                    break;
                                }
                                else
                                {
                                    measAfterHigh1 = Convert.ToBoolean(parsedStr[1]);
                                    measAfterHigh[0] = measAfterHigh1;
                                }                                    
                                break;
                            case "measAfterHigh2":
                                if (parsedStr[1].Length == 0)
                                {
                                    measAfterHigh2 = false;
                                    measAfterHigh[1] = measAfterHigh2;
                                    break;
                                }
                                else
                                {
                                    measAfterHigh2 = Convert.ToBoolean(parsedStr[1]);
                                    measAfterHigh[1] = measAfterHigh2;
                                }
                                break;
                            case "measAfterHigh3":
                                if (parsedStr[1].Length == 0)
                                {
                                    measAfterHigh3 = false;
                                    measAfterHigh[2] = measAfterHigh3;
                                    break;
                                }
                                else
                                {
                                    measAfterHigh3 = Convert.ToBoolean(parsedStr[1]);
                                    measAfterHigh[2] = measAfterHigh3;
                                }
                                break;
                            case "measAfterHigh4":
                                if (parsedStr[1].Length == 0)
                                {
                                    measAfterHigh4 = false;
                                    measAfterHigh[3] = measAfterHigh4;
                                    break;
                                }
                                else
                                {
                                    measAfterHigh4 = Convert.ToBoolean(parsedStr[1]);
                                    measAfterHigh[3] = measAfterHigh4;
                                }
                                break;
                            case "measAfterHigh5":
                                if (parsedStr[1].Length == 0)
                                {
                                    measAfterHigh5 = false;
                                    measAfterHigh[4] = measAfterHigh5;
                                    break;
                                }
                                else
                                {
                                    measAfterHigh5 = Convert.ToBoolean(parsedStr[1]);
                                    measAfterHigh[4] = measAfterHigh5;
                                }
                                break;
                            //---------------------------------------------------------------------------
                            case "measCP1":                           
                                if (parsedStr[1].Length == 0)
                                {
                                    measCP1 = false;
                                    measCP[0] = measCP1;
                                    break;
                                }
                                else
                                {
                                    measCP1 = Convert.ToBoolean(parsedStr[1]);
                                    measCP[0] = measCP1;
                                }                                    
                                break;
                            case "measCP2":
                                if (parsedStr[1].Length == 0)
                                {
                                    measCP2 = false;
                                    measCP[1] = measCP2;
                                    break;
                                }
                                else
                                {
                                    measCP2 = Convert.ToBoolean(parsedStr[1]);
                                    measCP[1] = measCP2;
                                }
                                break;
                            case "measCP3":
                                if (parsedStr[1].Length == 0)
                                {
                                    measCP3 = false;
                                    measCP[2] = measCP3;
                                    break;
                                }
                                else
                                {
                                    measCP3 = Convert.ToBoolean(parsedStr[1]);
                                    measCP[2] = measCP3;
                                }
                                break;
                            case "measCP4":
                                if (parsedStr[1].Length == 0)
                                {
                                    measCP4 = false;
                                    measCP[3] = measCP4;
                                    break;
                                }
                                else
                                {
                                    measCP4 = Convert.ToBoolean(parsedStr[1]);
                                    measCP[3] = measCP4;
                                }
                                break;
                            case "measCP5":
                                if (parsedStr[1].Length == 0)
                                {
                                    measCP5 = false;
                                    measCP[4] = measCP5;
                                    break;
                                }
                                else
                                {
                                    measCP5 = Convert.ToBoolean(parsedStr[1]);
                                    measCP[4] = measCP5;
                                }
                                break;
                            //----------------------------------------------------
                            case "cpFreqStart1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStart1);
                                    cpFreqStart[0] = cpFreqStart1;
                                }                                    
                                break;
                            case "cpFreqStart2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStart2);
                                    cpFreqStart[1] = cpFreqStart2;
                                }
                                break;
                            case "cpFreqStart3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStart3);
                                    cpFreqStart[2] = cpFreqStart3;
                                }
                                break;
                            case "cpFreqStart4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStart4);
                                    cpFreqStart[3] = cpFreqStart4;
                                }
                                break;
                            case "cpFreqStart5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStart5);
                                    cpFreqStart[4] = cpFreqStart5;
                                }
                                break;
                            //----------------------------------------------------
                            case "cpFreqStep1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStep1);
                                    cpFreqStep[0] = cpFreqStep1;
                                }                                    
                                break;
                            case "cpFreqStep2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStep2);
                                    cpFreqStep[1] = cpFreqStep2;
                                }
                                break;
                            case "cpFreqStep3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStep3);
                                    cpFreqStep[2] = cpFreqStep3;
                                }
                                break;
                            case "cpFreqStep4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStep4);
                                    cpFreqStep[3] = cpFreqStep4;
                                }
                                break;
                            case "cpFreqStep5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpFreqStep5);
                                    cpFreqStep[4] = cpFreqStep5;
                                }
                                break;
                            //----------------------------------------------------
                            case "cpNumSteps1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out cpNumSteps1);
                                    cpNumSteps[0] = cpNumSteps1;
                                }                                    
                                break;
                            case "cpNumSteps2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out cpNumSteps2);
                                    cpNumSteps[1] = cpNumSteps2;
                                }
                                break;
                            case "cpNumSteps3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out cpNumSteps3);
                                    cpNumSteps[2] = cpNumSteps3;
                                }
                                break;
                            case "cpNumSteps4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out cpNumSteps4);
                                    cpNumSteps[3] = cpNumSteps4;
                                }
                                break;
                            case "cpNumSteps5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    int.TryParse(parsedStr[1], out cpNumSteps5);
                                    cpNumSteps[4] = cpNumSteps5;
                                }
                                break;
                            //----------------------------------------------------
                            case "cpHigh1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpHigh1);
                                    cpHigh[0] = cpHigh1;
                                }                                    
                                break;
                            case "cpHigh2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpHigh2);
                                    cpHigh[1] = cpHigh2;
                                }
                                break;
                            case "cpHigh3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpHigh3);
                                    cpHigh[2] = cpHigh3;
                                }
                                break;
                            case "cpHigh4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpHigh4);
                                    cpHigh[3] = cpHigh4;
                                }
                                break;
                            case "cpHigh5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpHigh5);
                                    cpHigh[4] = cpHigh5;
                                }
                                break;
                            //----------------------------------------------------
                            case "cpLow1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpLow1);
                                    cpLow[0] = cpLow1;
                                }                                    
                                break;
                            case "cpLow2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpLow2);
                                    cpLow[1] = cpLow2;
                                }
                                break;
                            case "cpLow3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpLow3);
                                    cpLow[2] = cpLow3;
                                }
                                break;
                            case "cpLow4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpLow4);
                                    cpLow[3] = cpLow4;
                                }
                                break;
                            case "cpLow5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpLow5);
                                    cpLow[4] = cpLow5;
                                }
                                break;
                            //----------------------------------------------------
                            case "cpTrans1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpTrans1);
                                    cpTrans[0] = cpTrans1;
                                }                                    
                                break;
                            case "cpTrans2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpTrans2);
                                    cpTrans[1] = cpTrans2;
                                }
                                break;
                            case "cpTrans3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpTrans3);
                                    cpTrans[2] = cpTrans3;
                                }
                                break;
                            case "cpTrans4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpTrans4);
                                    cpTrans[3] = cpTrans4;
                                }
                                break;
                            case "cpTrans5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out cpTrans5);
                                    cpTrans[4] = cpTrans5;
                                }
                                break;
                            //----------------------------------------------------
                            case "numCycles":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                    int.TryParse(parsedStr[1], out numCycles);
                                break;
                            //----------------------------------------------------
                            case "measIV1":
                                if (parsedStr[1].Length == 0)
                                {
                                    measIV1 = false;
                                    measIV[0] = measIV1;
                                    break;
                                }
                                else
                                {
                                    measIV1 = Convert.ToBoolean(parsedStr[1]);
                                    measIV[0] = measIV1;
                                    if (measIV1)
                                    {
                                        measCP1 = false;     //charge pumping 
                                        measCP[0] = measCP1;
                                        measSpot1 = false;   //arbitrary spot meas
                                        measSpot[0] = measSpot1;
                                    }
                                    if (!measIV1)
                                    {
                                        measCP1 = false;    //charge pumping 
                                        measCP[0] = measCP1;
                                        measSpot1 = false;   //arbitrary spot meas
                                        measSpot[0] = measSpot1;
                                    }
                                }
                                break;
                            case "measIV2":
                                if (parsedStr[1].Length == 0)
                                {
                                    measIV2 = false;
                                    measIV[1] = measIV2;
                                    break;
                                }
                                else
                                {
                                    measIV2 = Convert.ToBoolean(parsedStr[1]);
                                    measIV[1] = measIV2;
                                    if (measIV2)
                                    {
                                        measCP2 = false;     //charge pumping 
                                        measCP[1] = measCP2;
                                        measSpot2 = false;   //arbitrary spot meas
                                        measSpot[1] = measSpot2;
                                    }
                                    if (!measIV2)
                                    {
                                        measCP2 = false;    //charge pumping 
                                        measCP[1] = measCP2;
                                        measSpot2 = false;   //arbitrary spot meas
                                        measSpot[1] = measSpot2;
                                    }
                                }
                                break;
                            case "measIV3":
                                if (parsedStr[1].Length == 0)
                                {
                                    measIV3 = false;
                                    measIV[2] = measIV3;
                                    break;
                                }
                                else
                                {
                                    measIV3 = Convert.ToBoolean(parsedStr[1]);
                                    measIV[2] = measIV3;
                                    if (measIV3)
                                    {
                                        measCP3 = false;     //charge pumping 
                                        measCP[2] = measCP3;
                                        measSpot3 = false;   //arbitrary spot meas
                                        measSpot[2] = measSpot3;
                                    }
                                    if (!measIV3)
                                    {
                                        measCP3 = false;    //charge pumping 
                                        measCP[2] = measCP3;
                                        measSpot3 = false;   //arbitrary spot meas
                                        measSpot[2] = measSpot3;
                                    }
                                }
                                break;
                            case "measIV4":
                                if (parsedStr[1].Length == 0)
                                {
                                    measIV4 = false;
                                    measIV[3] = measIV4;
                                    break;
                                }
                                else
                                {
                                    measIV4 = Convert.ToBoolean(parsedStr[1]);
                                    measIV[3] = measIV4;
                                    if (measIV4)
                                    {
                                        measCP4 = false;     //charge pumping 
                                        measCP[3] = measCP4;
                                        measSpot4 = false;   //arbitrary spot meas
                                        measSpot[3] = measSpot4;
                                    }
                                    if (!measIV4)
                                    {
                                        measCP4 = false;    //charge pumping 
                                        measCP[3] = measCP4;
                                        measSpot4 = false;   //arbitrary spot meas
                                        measSpot[3] = measSpot4;
                                    }
                                }
                                break;
                            case "measIV5":
                                if (parsedStr[1].Length == 0)
                                {
                                    measIV5 = false;
                                    measIV[4] = measIV5;
                                    break;
                                }
                                else
                                {
                                    measIV5 = Convert.ToBoolean(parsedStr[1]);
                                    measIV[4] = measIV5;
                                    if (measIV5)
                                    {
                                        measCP5 = false;     //charge pumping 
                                        measCP[4] = measCP5;
                                        measSpot5 = false;   //arbitrary spot meas
                                        measSpot[4] = measSpot5;
                                    }
                                    if (!measIV5)
                                    {
                                        measCP5 = false;    //charge pumping 
                                        measCP[4] = measCP5;
                                        measSpot5 = false;   //arbitrary spot meas
                                        measSpot[4] = measSpot5;
                                    }
                                }
                                break;
                            //----------------------------------------------------
                            case "IVGateStart1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStart1);
                                    ivGateStart[0] = ivGateStart1;
                                }                                    
                                break;
                            case "IVGateStart2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStart2);
                                    ivGateStart[1] = ivGateStart2;
                                }
                                break;
                            case "IVGateStart3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStart3);
                                    ivGateStart[2] = ivGateStart3;
                                }
                                break;
                            case "IVGateStart4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStart4);
                                    ivGateStart[3] = ivGateStart4;
                                }
                                break;
                            case "IVGateStart5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStart5);
                                    ivGateStart[4] = ivGateStart5;
                                }
                                break;
                            //----------------------------------------------------
                            case "IVGateStop1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStop1);
                                    ivGateStop[0] = ivGateStop1;
                                }                                    
                                break;
                            case "IVGateStop2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStop2);
                                    ivGateStop[1] = ivGateStop2;
                                }
                                break;
                            case "IVGateStop3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStop3);
                                    ivGateStop[2] = ivGateStop3;
                                }
                                break;
                            case "IVGateStop4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStop4);
                                    ivGateStop[3] = ivGateStop4;
                                }
                                break;
                            case "IVGateStop5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStop5);
                                    ivGateStop[4] = ivGateStop5;
                                }
                                break;
                            //----------------------------------------------------
                            case "IVGateStep1":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStep1);
                                    ivGateStep[0] = ivGateStep1;
                                }                                    
                                break;
                            case "IVGateStep2":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStep2);
                                    ivGateStep[1] = ivGateStep2;
                                }
                                break;
                            case "IVGateStep3":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStep3);
                                    ivGateStep[2] = ivGateStep3;
                                }
                                break;
                            case "IVGateStep4":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStep4);
                                    ivGateStep[3] = ivGateStep4;
                                }
                                break;
                            case "IVGateStep5":
                                if (parsedStr[1].Length == 0)
                                    break;
                                else
                                {
                                    double.TryParse(parsedStr[1], out ivGateStep5);
                                    ivGateStep[4] = ivGateStep5;
                                }
                                break;
                            //----------------------------------------------------
                            case "measSpot1":
                                if (parsedStr[1].Length == 0)
                                {
                                    measSpot1 = false;
                                    measSpot[0] = measSpot1;
                                    break;
                                }
                                else
                                {
                                    measSpot1 = Convert.ToBoolean(parsedStr[1]);
                                    measSpot[0] = measSpot1;
                                }                                    
                                break;
                            case "measSpot2":
                                if (parsedStr[1].Length == 0)
                                {
                                    measSpot2 = false;
                                    measSpot[1] = measSpot2;
                                    break;
                                }
                                else
                                {
                                    measSpot2 = Convert.ToBoolean(parsedStr[1]);
                                    measSpot[1] = measSpot2;
                                }
                                break;
                            case "measSpot3":
                                if (parsedStr[1].Length == 0)
                                {
                                    measSpot3 = false;
                                    measSpot[2] = measSpot3;
                                    break;
                                }
                                else
                                {
                                    measSpot3 = Convert.ToBoolean(parsedStr[1]);
                                    measSpot[2] = measSpot3;
                                }
                                break;
                            case "measSpot4":
                                if (parsedStr[1].Length == 0)
                                {
                                    measSpot4 = false;
                                    measSpot[3] = measSpot4;
                                    break;
                                }
                                else
                                {
                                    measSpot4 = Convert.ToBoolean(parsedStr[1]);
                                    measSpot[3] = measSpot4;
                                }
                                break;
                            case "measSpot5":
                                if (parsedStr[1].Length == 0)
                                {
                                    measSpot5 = false;
                                    measSpot[4] = measSpot5;
                                    break;
                                }
                                else
                                {
                                    measSpot5 = Convert.ToBoolean(parsedStr[1]);
                                    measSpot[4] = measSpot5;
                                }
                                break;
                            //----------------------------------------------------
                            case "pGate": // e.g. pGate=-1,-0.5,-0.3
                                string[] pGValues = parsedStr[1].Split(',');
                                for (int j = 0; j < pGValues.Length; j++)
                                {
                                    double pTemp = 0;
                                    double.TryParse(pGValues[j], out pTemp);
                                    pMeasGate.Add(pTemp);
                                }
                                break;
                            //----------------------------------------------------
                            case "pDrain": // e.g. pDrain=-0.05,-0.05,-0.7
                                string[] pDValues = parsedStr[1].Split(',');
                                for (int j = 0; j < pDValues.Length; j++)
                                {
                                    double pTemp = 0;
                                    double.TryParse(pDValues[j], out pTemp);
                                    pMeasDrain.Add(pTemp);
                                }
                                break;
                            default:
                                // exit program (should never get here)
                                throw new Exception();
                        } // end switch
                    } // end else
                }
            }
            catch (Exception)
            {
                // problem parsing the data
                Console.WriteLine("..........Error Parsing the input parameters!..............");
                Console.WriteLine("Error on " + tempParam);
                Console.ReadKey();
                return false;
            }

            for (int i = 0; i < TotNumDUTs; i++)
            {
                if(vGateStress[i] != 0.0 && vGateSense[i] != 0.0)
                {
                    if (vGateStress[i] < -1e9 || (!measCP[i] && !measIV[i] && !measSpot[i] && vGateSense[i] < -1e9))
                    {
                        // problem parsing the data
                        Console.WriteLine("..........Error with the input parameters!..............");
                        Console.WriteLine("........DUT#" + DutNum[i].ToString() + "........");
                        Console.WriteLine("vGateStress or vGateSense value not valid");
                        Console.ReadKey();
                        return false; // Don't have valid stress or sense voltages
                    }
                }
            }

            if (pMeasDrain.Count != pMeasGate.Count)
            {
                // problem parsing the data
                Console.WriteLine("..........Error Parsing the input parameters!..............");
                Console.WriteLine("Must have the same number of pGate and pDrain points");
                Console.ReadKey();
                return false;
            }
            return true;
        }

        static void writeParamsToConsole(int index)
        {
            Console.WriteLine("=========== Information of DUT#" + DutNum[index].ToString() + " ===========");
            Console.WriteLine(".......Channels Assigned to Gate, Drain, Source, and Substrate...........");
            Console.WriteLine("Channel# for Gate = " + channel1[index]);
            Console.WriteLine("Channel# for Drain = " + channel2[index]);
            Console.WriteLine("Channel# for Source = " + channel3[index]);
            Console.WriteLine("Channel# for Substrate = " + channel4[index]);
            Console.WriteLine("");
            Console.WriteLine("...............Stress Settings......................");
            Console.WriteLine("stressTime (sec):" + stressTime[index].ToString());
            Console.WriteLine("vGateStress (V):" + vGateStress[index].ToString());
            Console.WriteLine("vDrainStress (V):" + vDrainStress[index].ToString());
            if (acStress[index] == true)
            {
                Console.WriteLine("gateACLow (Low gate bias, V):" + vGateACLow[index].ToString());
                Console.WriteLine("drainACLow (Low drain bias, V):" + vDrainACLow[index].ToString());
                Console.WriteLine("freq (Frequency Hz):" + freq[index].ToString());
                Console.WriteLine("Duty Cycle %:" + dutyCycle[index].ToString());
                if (invStress[index] == true)
                {
                    Console.WriteLine("inverterStress (drain high while gate low) = true");
                }
                else
                {
                    Console.WriteLine("inverterStress (drain high while gate low) = false");
                }
            }
            Console.WriteLine("gateTransTime (gate pulse rise and fall time, sec):" + gateTransTime[index].ToString());
            Console.WriteLine("drainTransTime (drain pulse rise and fall time, sec):" + drainTransTime[index].ToString());
            Console.WriteLine("");
            Console.WriteLine("...............Relaxation Settings......................");
            Console.WriteLine("relaxTime:" + relaxTime[index].ToString());
            if (relaxTime[index] > 10e-9)
            {
                Console.WriteLine("vGateRelax (V):" + vGateRelax[index].ToString());
                Console.WriteLine("vDrainRelax (V):" + vDrainRelax[index].ToString());
            }
            Console.WriteLine("");
            Console.WriteLine("...............Measurement Settings......................");
            if (!measCP[index] && !measIV[index])
            {
                //Console.WriteLine("> Charge Pumping and/or IV Sweep");
                Console.WriteLine("> Regular spot measurement");    //modified by AK
                Console.WriteLine("vGateSense (V):" + vGateSense[index].ToString());
                Console.WriteLine("vDrainSense (V):" + vDrainSense[index].ToString());
                Console.WriteLine("measIRange (drain current measurement range, A):" + measIRange[index].ToString());
            }
            else if (measCP[index])
            {
                Console.WriteLine("> Measure Charge Pumping");
                Console.WriteLine("cpFreqStart (Hz):" + cpFreqStart[index].ToString());
                Console.WriteLine("cpFreqStep (Hz):" + cpFreqStep[index].ToString());
                Console.WriteLine("cpNumSteps:" + cpNumSteps[index].ToString());
                Console.WriteLine("cpHigh (high pulse voltage, V):" + cpHigh[index].ToString());
                Console.WriteLine("cpLow (low pulse voltage, V):" + cpLow[index].ToString());
                Console.WriteLine("cpTrans (rise time / fall time):" + cpTrans[index].ToString());
            }
            else if (measIV[index])
            {
                Console.WriteLine("> Measure IV Sweep");
                Console.WriteLine("ivGateStart (V):" + ivGateStart[index].ToString());
                Console.WriteLine("ivGateStop (V):" + ivGateStop[index].ToString());
                Console.WriteLine("ivGateStep (V):" + ivGateStep[index].ToString());
            }
            Console.WriteLine("measDelay (sec):" + measDelay[index].ToString());
            Console.WriteLine("measPoints:" + measPoints[index].ToString());
            Console.WriteLine("startAvgPoint:" + startAvgPoint[index].ToString());
            Console.WriteLine("stopAvgPoint:" + stopAvgPoint[index].ToString());
            Console.WriteLine("sampleInterval (sec):" + sampleInterval[index].ToString());
            double measurementWindow = measDelay[index] + measPoints[index] * sampleInterval[index];
            Console.WriteLine("Calculated Measurement Window (sec):" + measurementWindow.ToString());

            Console.WriteLine("initialSenseTime (sec):" + initialSenseTime[index].ToString());

            if (isLog[index])
            {
                Console.WriteLine("log option selected: measure logarithmically");
                Console.WriteLine("ppd (points per decade):" + ppd[index].ToString());
                Console.WriteLine("maxPulseWidth (max time allowed before making a measurement):" + maxPulseWidth[index].ToString());
            }
            else
            {
                Console.WriteLine("linear option selected: measure linear in time");
                Console.WriteLine("stepTime (linear meas time, sec):" + stepTime[index].ToString());
            }

            Console.WriteLine("gateTransTime (sec):" + gateTransTime[index].ToString());
            Console.WriteLine("drainTransTime (sec):" + drainTransTime[index].ToString());

            if (numCycles != 1)
            {
                Console.WriteLine("................Cycles to repeat stress/relaxation................");
                Console.WriteLine("numCycles:" + numCycles.ToString());
            }

            Console.WriteLine("");
        }

        static void takeData(int index)
        {         
            buildVectors(index);

            //--------------------------------------------------------------------
            WGFMU.setOperationMode(channel1[index], WGFMU.OPERATION_MODE_FASTIV);
            WGFMU.setOperationMode(channel2[index], WGFMU.OPERATION_MODE_FASTIV);

            //--------------------------------------------------------------------
            //Channel #1 measurement mode setting for IVSweep
            if (measIV[index]) WGFMU.setMeasureMode(channel1[index], WGFMU.MEASURE_MODE_VOLTAGE);

            //--------------------------------------------------------------------
            //Channel #2 current measurement mode setting --> Drain (or Source or Body)
            WGFMU.setMeasureMode(channel2[index], WGFMU.MEASURE_MODE_CURRENT);
            if (measIRange[index] <= 1e-6 * (1.001)) WGFMU.setMeasureCurrentRange(channel2[index], WGFMU.MEASURE_CURRENT_RANGE_1UA);
            else if (measIRange[index] <= 10e-6 * (1.001)) WGFMU.setMeasureCurrentRange(channel2[index], WGFMU.MEASURE_CURRENT_RANGE_10UA);
            else if (measIRange[index] <= 100e-6 * (1.001)) WGFMU.setMeasureCurrentRange(channel2[index], WGFMU.MEASURE_CURRENT_RANGE_100UA);
            else if (measIRange[index] <= 1e-3 * (1.001)) WGFMU.setMeasureCurrentRange(channel2[index], WGFMU.MEASURE_CURRENT_RANGE_1MA);
            else WGFMU.setMeasureCurrentRange(channel2[index], WGFMU.MEASURE_CURRENT_RANGE_10MA);

            //--------------------------------------------------------------------
            //Channel #1 voltage range setting --> Gate
            // set the gate range - may need to update with CP parameters (esp for low)
            double vRange = Math.Abs(vGateStress[index]);
            if (vRange < Math.Abs(vGateRelax[index])) vRange = Math.Abs(vGateRelax[index]);
            if (vRange < Math.Abs(vGateSense[index])) vRange = Math.Abs(vGateSense[index]);

            if (Math.Abs(vRange) <= 5 * 1.001) WGFMU.setMeasureVoltageRange(channel1[index], WGFMU.MEASURE_VOLTAGE_RANGE_5V);
            else if (Math.Abs(vRange) <= 10 * 1.001) WGFMU.setMeasureVoltageRange(channel1[index], WGFMU.MEASURE_VOLTAGE_RANGE_10V);

            //--------------------------------------------------------------------
            // Channel #2 voltage range setting
            // set the drain range -- check with ac vstressdrain low
            vRange = Math.Abs(vDrainStress[index]);
            if (vRange < Math.Abs(vDrainRelax[index])) vRange = Math.Abs(vDrainRelax[index]);
            if (vRange < Math.Abs(vDrainSense[index])) vRange = Math.Abs(vDrainSense[index]);

            if (Math.Abs(vRange) <= 5 * 1.001) WGFMU.setMeasureVoltageRange(channel2[index], WGFMU.MEASURE_VOLTAGE_RANGE_5V);
            else if (Math.Abs(vRange) <= 10 * 1.001) WGFMU.setMeasureVoltageRange(channel2[index], WGFMU.MEASURE_VOLTAGE_RANGE_10V);

            //--------------------------------------------------------------------
            // channel #3 for Source and Channel #4 dor Body if used
            if(channelSize[index] > 2)
            {
                WGFMU.setMeasureVoltageRange(channel3[index], WGFMU.MEASURE_VOLTAGE_RANGE_5V);

                if (channelSize[index] > 3)
                    WGFMU.setMeasureVoltageRange(channel4[index], WGFMU.MEASURE_VOLTAGE_RANGE_5V);
            }
        }

        private static DateTime startTime;
        private static DateTime endTime;

        private static void OnTimedEvent0(Object source, ElapsedEventArgs e)
        {
            TimeSpan elapsedTime = e.SignalTime - startTime;
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            if (elapsedTime.TotalSeconds < stressTime[0])
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = 0.0/{2:F1}", elapsedTime.TotalSeconds, stressTime[0], relaxTime[0]);
            }
            else
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = {2:F1}/{3:F1}", stressTime[0], stressTime[0], elapsedTime.TotalSeconds - stressTime[0], relaxTime[0]);
            }
            if (elapsedTime.TotalSeconds > stressTime[0] + relaxTime[0] + 20)
            {
                // if we go 20 seconds over then perhaps we have some errors that we need to deal with
                // Lets try to abort the test and write the results
                WGFMU.abort();

                writeResults(0, channel2[0], savePath[0], channel1[0], temp_stressTimePoint1, temp_relaxTimePoint1);
                Timer thisTimer = (Timer)source;
                thisTimer.Enabled = false;
            }
        }

        private static void OnTimedEvent1(Object source, ElapsedEventArgs e)
        {
            TimeSpan elapsedTime = e.SignalTime - startTime;
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            if (elapsedTime.TotalSeconds < stressTime[1])
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = 0.0/{2:F1}", elapsedTime.TotalSeconds, stressTime[1], relaxTime[1]);
            }
            else
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = {2:F1}/{3:F1}", stressTime[1], stressTime[1], elapsedTime.TotalSeconds - stressTime[1], relaxTime[1]);
            }
            if (elapsedTime.TotalSeconds > stressTime[1] + relaxTime[1] + 20)
            {
                // if we go 20 seconds over then perhaps we have some errors that we need to deal with
                // Lets try to abort the test and write the results
                WGFMU.abort();

                writeResults(1, channel2[1], savePath[1], channel1[1], temp_stressTimePoint2, temp_relaxTimePoint2);
                Timer thisTimer = (Timer)source;
                thisTimer.Enabled = false;
            }
        }

        private static void OnTimedEvent2(Object source, ElapsedEventArgs e)
        {
            TimeSpan elapsedTime = e.SignalTime - startTime;
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            if (elapsedTime.TotalSeconds < stressTime[2])
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = 0.0/{2:F1}", elapsedTime.TotalSeconds, stressTime[2], relaxTime[2]);
            }
            else
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = {2:F1}/{3:F1}", stressTime[2], stressTime[2], elapsedTime.TotalSeconds - stressTime[2], relaxTime[2]);
            }
            if (elapsedTime.TotalSeconds > stressTime[2] + relaxTime[2] + 20)
            {
                // if we go 20 seconds over then perhaps we have some errors that we need to deal with
                // Lets try to abort the test and write the results
                WGFMU.abort();

                writeResults(2, channel2[2], savePath[2], channel1[2], temp_stressTimePoint3, temp_relaxTimePoint3);               
                Timer thisTimer = (Timer)source;
                thisTimer.Enabled = false;
            }
        }

        private static void OnTimedEvent3(Object source, ElapsedEventArgs e)
        {
            TimeSpan elapsedTime = e.SignalTime - startTime;
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            if (elapsedTime.TotalSeconds < stressTime[3])
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = 0.0/{2:F1}", elapsedTime.TotalSeconds, stressTime[3], relaxTime[3]);
            }
            else
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = {2:F1}/{3:F1}", stressTime[3], stressTime[3], elapsedTime.TotalSeconds - stressTime[3], relaxTime[3]);
            }
            if (elapsedTime.TotalSeconds > stressTime[3] + relaxTime[3] + 20)
            {
                // if we go 20 seconds over then perhaps we have some errors that we need to deal with
                // Lets try to abort the test and write the results
                WGFMU.abort();

                writeResults(3, channel2[3], savePath[3], channel1[3], temp_stressTimePoint4, temp_relaxTimePoint4);
                Timer thisTimer = (Timer)source;
                thisTimer.Enabled = false;
            }
        }

        private static void OnTimedEvent4(Object source, ElapsedEventArgs e)
        {
            TimeSpan elapsedTime = e.SignalTime - startTime;
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            if (elapsedTime.TotalSeconds < stressTime[4])
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = 0.0/{2:F1}", elapsedTime.TotalSeconds, stressTime[4], relaxTime[4]);
            }
            else
            {
                Console.WriteLine("Stress Time = {0:F1}/{1:F1}; Relax Time = {2:F1}/{3:F1}", stressTime[4], stressTime[4], elapsedTime.TotalSeconds - stressTime[4], relaxTime[4]);
            }
            if (elapsedTime.TotalSeconds > stressTime[4] + relaxTime[4] + 20)
            {
                // if we go 20 seconds over then perhaps we have some errors that we need to deal with
                // Lets try to abort the test and write the results
                WGFMU.abort();

                writeResults(4, channel2[4], savePath[4], channel1[4], temp_stressTimePoint5, temp_relaxTimePoint5);
                Timer thisTimer = (Timer)source;
                thisTimer.Enabled = false;
            }
        }

        static void writeResults(int index, int channel, string filePath, int channelGate, List<double> stressTimePoint, List<double> relaxTimePoint)
        {
            //Some info on when the time of the measurement is returned
            //average
            // : 
            //Averaging time, in second. Numeric. 0 (no averaging), or 10-8 (10 ns) to 0.020971512 (approximately 20 ms), in 10-8 (10 ns) resolution. Do 
            //not have to exceed the interval value. If nonzero value is specified, the channel repeats measurement in 5 ns interval while the average
            //period1, and returns the averaging result data. For example, if a measurement starts at 0 ns and average=20 ns, 
            //measurement is performed at 0, 5, 10, and 15 ns. And time data for the averaging result data is 10 ns = (0+20)/2.

            List<string> lines = new List<string>();

            lines.Add("*BTITest=CVS");
            lines.Add("***********************************************");
            lines.Add("********** DUT#" + DutNum[index].ToString() + " Measurement **********");
            lines.Add("********** Stress Option **********");

            if (acStress[index] == false)
                lines.Add("*DC = true");
            else
                lines.Add("*AC = true");

            lines.Add("*vGateStress = " + vGateStress[index].ToString());
            lines.Add("*vDrainStress = " + vDrainStress[index].ToString());
            lines.Add("*stressTime = " + stressTime[index].ToString());
            lines.Add("*vGateRelax = " + vGateRelax[index].ToString());
            lines.Add("*vDrainRelax = " + vDrainRelax[index].ToString());
            lines.Add("*relaxTime = " + relaxTime[index].ToString());
            lines.Add("*gateTransTime = " + gateTransTime[index].ToString());
            lines.Add("*drainTransTime = " + drainTransTime[index].ToString());

            if (acStress[index] == true)
            {
                lines.Add("*vGateACLow = " + vGateACLow[index].ToString());
                lines.Add("*vDrainACLow = " + vDrainACLow[index].ToString());
                lines.Add("*Frequency = " + freq[index].ToString());
                lines.Add("*dutyCycle = " + dutyCycle[index].ToString());
                lines.Add("*skew =" + skew[index].ToString());

                if (invStress[index] == true)
                    lines.Add("*InverterStress = true");
                else
                    lines.Add("*InverterStress = false");
            }

            lines.Add("***********************************************");
            lines.Add("********** Measurement Option **********");

            if (measCP[index] == true)
                lines.Add("*Charge Pumping = true");
            else if (measIV[index] == true)
                lines.Add("*IV Sweep = true");
            else
            {
                lines.Add("*Spot Measurement = true");
                lines.Add("*vGateSense = " + vGateSense[index].ToString());
            }

            lines.Add("*vDrainSense = " + vDrainSense[index].ToString());
            lines.Add("*measIRange = " + measIRange[index].ToString());
            lines.Add("*initialSenseTime = " + initialSenseTime[index].ToString());
            lines.Add("*measDelay = " + measDelay[index].ToString());
            lines.Add("*measPoints = " + measPoints[index].ToString());
            lines.Add("*startAvgPoint = " + startAvgPoint[index].ToString());
            lines.Add("*stopAvgPoint = " + stopAvgPoint[index].ToString());
            lines.Add("*sampleInterval = " + sampleInterval[index].ToString());

            if (measCP[index])
            {
                lines.Add("*cpFreqStart = " + cpFreqStart[index].ToString());
                lines.Add("*cpFreqStep = " + cpFreqStep[index].ToString());
                lines.Add("*cpNumSteps = " + cpNumSteps[index].ToString());
                lines.Add("*CPHighVoltage = " + cpHigh[index].ToString());
                lines.Add("*CPLowVoltage = " + cpLow[index].ToString());
                lines.Add("*CPtrtf = " + cpTrans[index].ToString());
            }
            else if (measIV[index])
            {
                lines.Add("*IVGateVStart = " + ivGateStart[index].ToString());
                lines.Add("*IVGateVStop = " + ivGateStop[index].ToString());
                lines.Add("*IVGateVStep = " + ivGateStep[index].ToString());
            }

            lines.Add("***********************************************");
            lines.Add("********** Measurement Time Interval **********");

            if (isLog[index] == true)
            {
                lines.Add("*Log = true");
                lines.Add("*ppd = " + ppd[index].ToString());
            }

            else
            {
                lines.Add("*Linear = true");
                lines.Add("*stepTime = " + stepTime[index].ToString());
            }

            lines.Add("***********************************************");
            lines.Add("*StartTime =" + startTime.ToString());
            lines.Add("*EndTime =" + endTime.ToString());
            lines.Add("***********************************************");

            // add the header to the data (what each column is)
            // Columns we want: Time, ACStressTimeRaw, StressTimeRaw, RelaxTimeRaw, MeasurePoint, MeasureTimeRaw, IDrain, ACStressTime, stressTime1, StressIDrain, relaxTime1, RelaxIDrain
            if (acStress[index] == true && relaxTime[index] > 0 && !measCP[index] && !measIV[index] && !measSpot[index])
            {
                lines.Add("Time\tACStressTimeRaw\tStressTimeRaw\tRelaxTimeRaw\tMeasurePoint\tMeasureTimeRaw\tIDrain\tACStressTime\tStressTime\tStressIDrain\tRelaxTime\tRelaxIDrain");
            }
            else if (acStress[index] == true && relaxTime[index] > 0 && measCP[index])
            {
                lines.Add("Time\tACStressTimeRaw\tStressTimeRaw\tRelaxTimeRaw\tMeasurePoint\tCPFreq\tMeasureTimeRaw\tIcp\tACStressTime\tStressTime\tStressCPFreq\tStressIcp\tRelaxTime\tRelaxCPFreq\tRelaxIcp");
            }
            else if (acStress[index] == true && relaxTime[index] > 0 && measIV[index])
            {
                lines.Add("Time\tACStressTimeRaw\tStressTimeRaw\tRelaxTimeRaw\tMeasurePoint\tVGate\tMeasureTimeRaw\tIDrain\tACStressTime\tStressTime\tStressVGate\tStressIDrain\tRelaxTime\tRelaxVGate\tRelaxIDrain");
            }
            else if (acStress[index] == true && relaxTime[index] > 0 && measSpot[index])
            {
                lines.Add("Time\tACStressTimeRaw\tStressTimeRaw\tRelaxTimeRaw\tMeasurePoint\tVGate\tVDrain\tMeasureTimeRaw\tIDrain\tACStressTime\tStressTime\tStressVGate\tStressVDrain\tStressIDrain\tRelaxTime\tRelaxVGate\tRelaxVDrain\tRelaxIDrain");
            }
            // acStress[index] no relaxation
            else if (acStress[index] == true && !measCP[index] && !measIV[index] && !measSpot[index])
            {
                lines.Add("Time\tACStressTimeRaw\tStressTimeRaw\tMeasurePoint\tMeasureTimeRaw\tIDrain\tACStressTime\tStressTime\tStressIDrain");
            }
            else if (acStress[index] == true && measCP[index])
            {
                lines.Add("Time\tACStressTimeRaw\tStressTimeRaw\tMeasurePoint\tCPFreq\tMeasureTimeRaw\tIcp\tACStressTime\tStressTime\tStressCPFreq\tStressIcp");
            }
            else if (acStress[index] == true && measIV[index])
            {
                lines.Add("Time\tACStressTimeRaw\tStressTimeRaw\tMeasurePoint\tVGate\tMeasureTimeRaw\tIDrain\tACStressTime\tStressTime\tStressVGate\tStressIDrain");
            }
            else if (acStress[index] == true && measSpot[index])
            {
                lines.Add("Time\tACStressTimeRaw\tStressTimeRaw\tMeasurePoint\tVGate\tVDrain\tMeasureTimeRaw\tIDrain\tACStressTime\tStressTime\tStressVGate\tStressVDrain\tStressIDrain");
            }
            // no acStress[index] with relaxation
            else if (relaxTime[index] > 0 && !measCP[index] && !measIV[index] && !measSpot[index])
            {
                lines.Add("Time\tStressTimeRaw\tRelaxTimeRaw\tMeasurePoint\tMeasureTimeRaw\tIDrain\tStressTime\tStressIDrain\tRelaxTime\tRelaxIDrain");
            }
            else if (relaxTime[index] > 0 && measCP[index])
            {
                lines.Add("Time\tStressTimeRaw\tRelaxTimeRaw\tMeasurePoint\tCPFreq\tMeasureTimeRaw\tIcp\tStressTime\tStressCPFreq\tStressIcp\tRelaxTime\tRelaxCPFreq\tRelaxIcp");
            }
            else if (relaxTime[index] > 0 && measIV[index])
            {
                lines.Add("Time\tStressTimeRaw\tRelaxTimeRaw\tMeasurePoint\tVGate\tMeasureTimeRaw\tIDrain\tStressTime\tStressVGate\tStressIDrain\tRelaxTime\tRelaxVGate\tRelaxIDrain");
            }
            else if (relaxTime[index] > 0 && measSpot[index])
            {
                lines.Add("Time\tStressTimeRaw\tRelaxTimeRaw\tMeasurePoint\tVGate\tVDrain\tMeasureTimeRaw\tIDrain\tStressTime\tStressVGate\tStressIDrain\tRelaxTime\tRelaxVGate\tRelaxVDrain\tRelaxIDrain");
            }
            // no acStress[index] no relaxation
            else if (!measCP[index] && !measIV[index])
            {
                lines.Add("Time\tStressTimeRaw\tMeasurePoint\tMeasureTimeRaw\tIDrain\tStressTime\tStressIDrain");
            }
            else if (measCP[index])
            {
                lines.Add("Time\tStressTimeRaw\tMeasurePoint\tCPFreq\tMeasureTimeRaw\tStressIcp\tStressTime\tStressCPFreq\tStressIcp");
            }
            else if (measIV[index])
            {
                lines.Add("Time\tStressTimeRaw\tMeasurePoint\tVGate\tMeasureTimeRaw\tStressIDrain\tStressTime\tStressVGate\tStressIDrain");
            }
            else if (measSpot[index])
            {
                lines.Add("Time\tStressTimeRaw\tMeasurePoint\tVGate\tVDrain\tMeasureTimeRaw\tStressIDrain\tStressTime\tStressVGate\tStressVDrain\tStressIDrain");
            }

            // Vectors for time and Idrain
            int measuredSize = 1; // holds all the measured data for the measured current
            int total = 1;
            WGFMU.getMeasureValueSize(channel, ref measuredSize, ref total);
            List<double> times = new List<double>();
            List<double> values = new List<double>();

            int measuredSizeGate = 1; // holds all the measured voltage data on the gate
            int totalGate = 1;
            WGFMU.getMeasureValueSize(channelGate, ref measuredSizeGate, ref totalGate);
            List<double> timesGate = new List<double>();
            List<double> valuesGate = new List<double>();

            if (measuredSize > 0) // no data measured
            {
                // fetch the data for the current
                for (int i = 0; i < measuredSize; i++)
                {
                    double time = 0;
                    double value = 0;
                    WGFMU.getMeasureValue(channel, i, ref time, ref value);
                    times.Add(time);
                    values.Add(value);
                }
                // fetch the data for the gate voltage of the IV sweep if performed
                if (measIV[index] || measSpot[index])
                {
                    for (int i = 0; i < measuredSizeGate; i++)
                    {
                        double time = 0;
                        double value = 0;
                        WGFMU.getMeasureValue(channelGate, i, ref time, ref value);
                        timesGate.Add(time);
                        valuesGate.Add(value);
                    }
                }

                List<double> CPFreqRaw = new List<double>(); // raw value

                if (measCP[index])
                {
                    for (int i = 0; i < measuredSize;)
                    {
                        for (int l = 0; l < cpNumSteps[index]; l++)
                        {
                            for (int j = 0; j < measPoints[index]; j++)
                            {
                                double tempFreq = cpFreqStart[index] + cpFreqStep[index] * l;
                                CPFreqRaw.Add(tempFreq);
                            }
                            i += measPoints[index];
                        }
                    }
                }

                List<double> VDrainRaw = new List<double>(); // raw VDrain value

                if (measSpot[index])
                {
                    for (int i = 0; i < measuredSize;)
                    {
                        for (int l = 0; l < pMeasGate.Count; l++)
                        {
                            for (int j = 0; j < measPoints[index]; j++)
                            {
                                double tempVDrain = pMeasDrain[l];
                                VDrainRaw.Add(tempVDrain);
                            }
                            i += measPoints[index];
                        }
                    }
                }

                // already have our vectors of stress and relax. Need vector of IDrain (Icp, etc)
                List<double> IDrainPoint = new List<double>(); // or Icp - averaged value over measurement window
                List<double> stressIDrainPoint = new List<double>();
                List<double> relaxIDrainPoint = new List<double>();

                // For when we measure the gate voltage
                List<double> VGatePoint = new List<double>(); // averaged value over measurement window
                List<double> stressVGatePoint = new List<double>();
                List<double> relaxVGatePoint = new List<double>();

                // Create the CP freq list
                List<double> CPFreq = new List<double>(); // value of the measurement window
                List<double> stressCPFreq = new List<double>();
                List<double> relaxCPFreq = new List<double>();

                // Create the drain voltage
                List<double> VDrainPoint = new List<double>(); // average value over measurement window
                List<double> stressVDrainPoint = new List<double>();
                List<double> relaxVDrainPoint = new List<double>();

                // ok, calc the average IDrain/Icp for each point
                for (int i = 0; i < measuredSize;)
                {
                    double avgValue = 0;
                    for (int j = startAvgPoint[index]; j < stopAvgPoint[index] + 1; j++)
                    {
                        avgValue += values[j + i] / (stopAvgPoint[index] - startAvgPoint[index] + 1);
                    }
                    IDrainPoint.Add(avgValue);
                    i += measPoints[index];
                }
                // calc the average VGate for each point
                if (measIV[index] || measSpot[index])
                {
                    for (int i = 0; i < measuredSizeGate;)
                    {
                        double avgValue = 0;
                        for (int j = startAvgPoint[index]; j < stopAvgPoint[index] + 1; j++)
                        {
                            avgValue += valuesGate[j + i] / (stopAvgPoint[index] - startAvgPoint[index] + 1);
                        }
                        VGatePoint.Add(avgValue);
                        i += measPoints[index];
                    }
                }
                if (measCP[index])
                {
                    // calc the value for each measurement window (average)
                    for (int i = 0; i < measuredSize / measPoints[index];)
                    {
                        for (int l = 0; l < cpNumSteps[index]; l++)
                        {
                            double tempFreq = cpFreqStart[index] + cpFreqStep[index] * l;
                            CPFreq.Add(tempFreq);
                        }
                        i += cpNumSteps[index];
                    }
                }
                if (measSpot[index])
                {
                    // calc the value for each measurement window (average)
                    for (int i = 0; i < measuredSize / measPoints[index];)
                    {
                        for (int l = 0; l < pMeasGate.Count; l++)
                        {
                            double tempValue = pMeasDrain[l];
                            VDrainPoint.Add(tempValue);
                        }
                        i += cpNumSteps[index];
                    }
                }

                // add in a zero to our stress time (for the initial point measured) t0
                stressTimePoint.Insert(0, 0.0);
                relaxTimePoint.Insert(0, 0.0);

                // update the stressTimePoint and relaxTimePoint by the number of measurement windows
                for (int i = stressTimePoint.Count - 1; i >= 0; i--)
                {
                    double tempTime = stressTimePoint[i];
                    for (int l = 1; l < measWindowsPerSense[index]; l++)
                    {
                        stressTimePoint.Insert(i, tempTime);
                    }
                }
                for (int i = relaxTimePoint.Count - 1; i >= 0; i--)
                {
                    double tempTime = relaxTimePoint[i];
                    for (int l = 1; l < measWindowsPerSense[index]; l++)
                    {
                        relaxTimePoint.Insert(i, tempTime);
                    }
                }

                // now separate into stress and relax points
                for (int i = 0; i < stressTimePoint.Count; i++)
                {
                    stressIDrainPoint.Add(IDrainPoint[i]);
                }
                for (int i = 0; i < relaxTimePoint.Count; i++)
                {
                    relaxIDrainPoint.Add(IDrainPoint[i + stressIDrainPoint.Count - measWindowsPerSense[index]]);
                }
                // now do the same for the gate voltage
                if (measIV[index] || measSpot[index])
                {
                    for (int i = 0; i < stressTimePoint.Count; i++)
                    {
                        stressVGatePoint.Add(VGatePoint[i]);
                    }
                    for (int i = 0; i < relaxTimePoint.Count; i++)
                    {
                        relaxVGatePoint.Add(VGatePoint[i + stressVGatePoint.Count - measWindowsPerSense[index]]);
                    }
                }
                if (measCP[index])
                {
                    for (int i = 0; i < stressTimePoint.Count; i++)
                    {
                        stressCPFreq.Add(CPFreq[i]);
                    }
                    for (int i = 0; i < relaxTimePoint.Count; i++)
                    {
                        relaxCPFreq.Add(CPFreq[i + stressCPFreq.Count - measWindowsPerSense[index]]);
                    }
                }
                if (measSpot[index])
                {
                    for (int i = 0; i < stressTimePoint.Count; i++)
                    {
                        stressVDrainPoint.Add(VDrainPoint[i]);
                    }
                    for (int i = 0; i < relaxTimePoint.Count; i++)
                    {
                        relaxVDrainPoint.Add(VDrainPoint[i + stressVDrainPoint.Count - measWindowsPerSense[index]]);
                    }
                }

                int k = 0;
                for (int i = 0; i < times.Count; i++)
                {
                    // now we write all the vectors out
                    // Main time vector
                    string line = "";
                    line += times[i].ToString();                                // Time

                    // Stress Vectors
                    if (acStress[index] == true && k < stressTimePoint.Count)
                    {
                        line += "\t" + stressTimePoint[k].ToString();           // ACStressTimeRaw
                        double temp = stressTimePoint[k] * dutyCycle[index] * 0.01;
                        line += "\t" + temp.ToString();                         // StressTimeRaw
                    }
                    else if (acStress[index] == true) line += "\t\t";                  // blank writes for when in relax part
                    else if (acStress[index] == false && k < stressTimePoint.Count)    // StressTimeRaw
                    {
                        line += "\t" + stressTimePoint[k].ToString();
                    }
                    else line += "\t";                                          // blank writes for when in relax part

                    // Relax Vectors
                    if (relaxTime[index] > 0)
                    {
                        if (k >= (stressTimePoint.Count - 1))
                        {
                            line += "\t" + relaxTimePoint[k - stressTimePoint.Count + 1].ToString();    // RelaxTimeRaw
                        }
                        else line += "\t";
                    }

                    // Sample Number
                    int result;
                    Math.DivRem(i + 1, measPoints[index], out result);

                    if (result == 0) result = measPoints[index];
                    line += "\t" + result.ToString();             // Measurement Sample

                    if (measCP[index])                 // CP freq
                    {
                        line += "\t" + CPFreqRaw[i].ToString();
                    }
                    if (measIV[index] || measSpot[index])                  // IV VGate
                    {
                        line += "\t" + valuesGate[i].ToString();
                        if (measSpot[index])            // Vdrain changes
                        {
                            line += "\t" + VDrainRaw[i].ToString();
                        }
                    }

                    double temp2;
                    if (i > measPoints[index] - 1) temp2 = times[i - k * measPoints[index]];
                    else temp2 = times[i];

                    line += "\t" + temp2.ToString();                                                // Sample Measurement Time
                    line += "\t" + values[i].ToString();                                            // IDrain

                    // Now the Stress columns

                    if (acStress[index] == true && i < stressTimePoint.Count)
                    {
                        line += "\t" + stressTimePoint[i].ToString();           // ACStressTime
                        double temp = stressTimePoint[i] * dutyCycle[index] * 0.01;
                        line += "\t" + temp.ToString();                         // stressTime
                    }
                    else if (acStress[index] == true) line += "\t\t";                  // Blank ACStressTime and stressTime
                    else if (acStress[index] == false && i < stressTimePoint.Count)
                    {
                        line += "\t" + stressTimePoint[i].ToString();           // stressTime (DC)
                    }
                    else line += "\t";                                          // Blank stressTime (DC)
                    if (i < stressTimePoint.Count)
                    {
                        if (measCP[index]) line += "\t" + stressCPFreq[i].ToString();
                        if (measIV[index] || measSpot[index]) line += "\t" + stressVGatePoint[i].ToString();
                        if (measSpot[index]) line += "\t" + stressVDrainPoint[i].ToString();
                        line += "\t" + stressIDrainPoint[i].ToString();
                    }
                    else
                    {
                        if (measCP[index] || measIV[index]) line += "\t";
                        if (measSpot[index]) line += "\t\t";
                        line += "\t";
                    }

                    // Now the Relax columns
                    if (relaxTime[index] > 0)
                    {
                        if (i < relaxTimePoint.Count)
                        {
                            line += "\t" + relaxTimePoint[i].ToString();
                            if (measCP[index]) line += "\t" + relaxCPFreq[i].ToString();
                            if (measIV[index] || measSpot[index]) line += "\t" + relaxVGatePoint[i].ToString();
                            if (measSpot[index]) line += "\t" + relaxVDrainPoint[i].ToString();
                            line += "\t" + relaxIDrainPoint[i].ToString();
                        }
                    }

                    if (result == measPoints[index] && i != 0) k++;

                    lines.Add(line); // add the line to our list of lines to record
                }

            }
            else
            {
                lines.Add("Error during measurement");
            }
            System.IO.File.WriteAllLines(@filePath, lines.ToArray());
        }

        /// <summary>
        /// Builds the stress/relaxation measurement sequence. In this version all the steps, stress/relaxation, 
        /// sense are made into sequences which are then strung together to create the final measurement sequence
        /// </summary>
        static void buildVectors(int index)
        {            
            stressTimePoint = new List<double>();
            relaxTimePoint = new List<double>();
            List<double> eventPoint = new List<double>();

            if (acStress[index]) period[index] = 1 / freq[index];
            else period[index] = 1e-6; // rounded to the 1us

            double periodRelax = 1e-6;

            // calculate the time at which to interupt the the stress / relaxation for a measurement
            if (isLog[index]) // log interupation
            {
                // calculate where we interupt the stress
                double multiFactor = Math.Pow(10, 1.0 / ppd[index]);

                if (!acStress[index])
                    period[index] = 1e-6; // rounded to the 1us

                // stress part
                // number of dec.
                double dec = Math.Log10(stressTime[index] / initialSenseTime[index]);
                if (initialSenseTime[index] < period[index]) dec = Math.Log10(stressTime[index] / period[index]); // if the user has a initial sense time less than 1 pulse make equal to one pulse

                int guessNumPoints = (int)(dec * ppd[index]);

                if (initialSenseTime[index] < period[index])
                {
                    if (stressTime[index] == 0 && relaxTime[index] == 0) ; //done nothing
                    else stressTimePoint.Add(period[index]);
                }
                    
                else
                {
                    if (stressTime[index] == 0 && relaxTime[index] == 0) ; //done nothing
                    else stressTimePoint.Add(initialSenseTime[index]);
                }
                    
                for (int i = 0; i < guessNumPoints + 100; i++)
                {
                    if (stressTimePoint[i] * multiFactor < stressTime[index])
                    {
                        stressTimePoint.Add(stressTimePoint[i] * multiFactor);
                    }
                    else
                    {
                        if (stressTimePoint[i] < stressTime[index] * (0.99))
                        {
                            stressTimePoint.Add(stressTime[index]);
                        }
                        break;
                    }
                }

                // clean up
                for (int i = stressTimePoint.Count - 1; i > 0; i--)
                {
                    if (stressTimePoint[i] - stressTimePoint[i - 1] < period[index])
                    {
                        stressTimePoint.RemoveAt(i);
                    }
                }

                // clean up using the max pulse width
                for (int i = 1; i < stressTimePoint.Count; i++)
                {
                    if (stressTimePoint[i] - stressTimePoint[i - 1] > maxPulseWidth[index])
                    {
                        // Then remove all points between this and the last
                        for (int j = stressTimePoint.Count - 2; j >= i; j--)
                        {
                            stressTimePoint.RemoveAt(j);
                        }

                        stressTimePoint.Insert(i, stressTimePoint[i - 1] + maxPulseWidth[index]);

                        // Now insert points at the max pulse width until we reach the end
                        for (int j = i + 1; j < stressTimePoint.Count; j++)
                        {
                            if (stressTimePoint[j] - stressTimePoint[j - 1] > 3600)
                            {
                                stressTimePoint.Insert(j, stressTimePoint[j - 1] + maxPulseWidth[index]);
                            }
                        }
                        // Exit the for loop
                        break;
                    }
                }

                // relaxation part
                // number of dec.
                dec = Math.Log10(relaxTime[index] / initialSenseTime[index]);

                guessNumPoints = (int)(dec * ppd[index]);
                if (relaxTime[index] > 0)
                {
                    relaxTimePoint.Add(initialSenseTime[index]);
                    for (int i = 0; i < guessNumPoints + 100; i++)
                    {
                        if (relaxTimePoint[i] * multiFactor < relaxTime[index])
                        {
                            relaxTimePoint.Add(Math.Round(relaxTimePoint[i] * multiFactor, 8));
                        }
                        else
                        {
                            if (relaxTimePoint[i] < relaxTime[index] * (0.99))
                            {
                                relaxTimePoint.Add(relaxTime[index]);
                            }
                            break;
                        }
                    }
                }

                if (relaxTime[index] > 0)
                {
                    for (int i = relaxTimePoint.Count - 1; i > 0; i--)
                    {
                        if (relaxTimePoint[i] - relaxTimePoint[i - 1] < periodRelax)
                        {
                            relaxTimePoint.RemoveAt(i);
                        }
                    }

                    // clean up using the max pulse width
                    for (int i = 1; i < relaxTimePoint.Count; i++)
                    {
                        if (relaxTimePoint[i] - relaxTimePoint[i - 1] > maxPulseWidth[index])
                        {
                            // Then remove all points between this and the last
                            for (int j = relaxTimePoint.Count - 2; j >= i; j--)
                            {
                                relaxTimePoint.RemoveAt(j);
                            }

                            relaxTimePoint.Insert(i, relaxTimePoint[i - 1] + maxPulseWidth[index]);

                            // Now insert points at the max pulse width until we reach the end
                            for (int j = i + 1; j < relaxTimePoint.Count; j++)
                            {
                                if (relaxTimePoint[j] - relaxTimePoint[j - 1] > 3600)
                                {
                                    relaxTimePoint.Insert(j, relaxTimePoint[j - 1] + maxPulseWidth[index]);
                                }
                            }
                            // Exit the for loop
                            break;
                        }
                    }
                }
            }
            else // linear stress
            {
                int numOfSteps = (int)((stressTime[index] - initialSenseTime[index]) / stepTime[index]) + 1;

                if (stressTime[index] == 0 && relaxTime[index] == 0) ; //done nothing
                else stressTimePoint.Add(initialSenseTime[index]);
               
                for (int i = 1; i <= numOfSteps; i++)
                {
                    stressTimePoint.Add(stepTime[index] * i + initialSenseTime[index]);
                }

                if (relaxTime[index] > 0)
                {
                    numOfSteps = (int)((relaxTime[index] - initialSenseTime[index]) / stepTime[index]) + 1;
                    relaxTimePoint.Add(initialSenseTime[index]);
                    for (int i = 1; i <= numOfSteps; i++)
                    {
                        relaxTimePoint.Add(stepTime[index] * i + initialSenseTime[index]);
                    }
                }
            }

            // Flip the drain high and low voltage if we are doing inverter stress
            if (invStress[index] == true)
            {
                double temp = vDrainStress[index];
                vDrainStress[index] = vDrainACLow[index];
                vDrainACLow[index] = temp;
            }

            // calc meas window
            double measWindow = sampleInterval[index] * measPoints[index] + measDelay[index];

            int outMode = 12000;

            // Build the measurement pattern.
            // Measurement Sequence is as follows
            // 1) Initial Measurement
            // 2) Bring to stress voltage (DC or AC condition) from sense condition

            // Start of cycle slope

            // 3) Stress / Sense / Stress / Sense (end on sense part)
            // 4) Relax / Sense / Relax / Sense (end on sense part)

            // loop cycle

            // for these measurments should have the following transistions
            // a) zero to measure condition
            // b) stress to sense condition
            // c) sense to measure condition
            // and the following cycles
            // Measure Cycle
            // Stress Cycle
            // Relax Cycle

            // Perform all the calculations necessary for Frequency and Duty Cycle so we use the right timing

            gateHighTime[index] = dutyCycle[index] * 0.01 * period[index] - gateTransTime[index];
            gateLowTime[index] = (100 - dutyCycle[index]) * 0.01 * period[index] - gateTransTime[index];
            drainHighTime[index] = dutyCycle[index] * 0.01 * period[index] - drainTransTime[index];// -  2*skew1;
            drainLowTime[index] = (100 - dutyCycle[index]) * 0.01 * period[index] - drainTransTime[index];//+ 2*skew1;

            // correct the step size
            if (ivGateStart[index] > ivGateStop[index]) ivGateStep[index] = -1.0 * Math.Abs(ivGateStep[index]);
            else ivGateStep[index] = Math.Abs(ivGateStep[index]);
            int tempNumSteps = (int)((ivGateStop[index] - ivGateStart[index]) / (ivGateStep[index]) + 1);
            // Build the measurement pattern
            if (!measCP[index]) // regular DC measurement of the Drain current
            {
                // create measurement pattern before stress
                WGFMU.createPattern("GATE_MEAS_PTN" + DutNum[index].ToString(), 0);     // start at the end of the off cycle of the stress pulse
                WGFMU.createPattern("DRAIN_MEAS_PTN" + DutNum[index].ToString(), 0);   // start at the end of the off cycle of the stress pulse

                if (channelSize[index] > 2)
                {
                    WGFMU.createPattern("SOURCE_MEAS_PTN" + DutNum[index].ToString(), 0);

                    if(channelSize[index] > 3)
                        WGFMU.createPattern("SUB_MEAS_PTN" + DutNum[index].ToString(), 0);
                }

                if (measIV[index])
                {
                    measWindowsPerSense[index] = tempNumSteps;
                    for (int i = 0; i < tempNumSteps; i++)
                    {
                        WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], ivGateStart[index] + i * ivGateStep[index]);
                        WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), measWindow, ivGateStart[index] + i * ivGateStep[index]);
                        WGFMU.setMeasureEvent("GATE_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        WGFMU.addVector("DRAIN_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], vDrainSense[index]);
                        WGFMU.addVector("DRAIN_MEAS_PTN" + DutNum[index].ToString(), measWindow, vDrainSense[index]);
                        WGFMU.setMeasureEvent("DRAIN_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if (channelSize[index] > 2)
                        {
                            WGFMU.addVector("SOURCE_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                            WGFMU.addVector("SOURCE_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                            WGFMU.setMeasureEvent("SOURCE_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                                measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if(channelSize[index] > 3)
                            {
                                WGFMU.addVector("SUB_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                                WGFMU.addVector("SUB_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                                WGFMU.setMeasureEvent("SUB_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                                    measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            }                            
                        }
                    }
                }
                if (measSpot[index])
                {
                    for (int i = 0; i < pMeasGate.Count; i++)
                    {
                        WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], pMeasGate[i]);
                        WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), measWindow, pMeasGate[i]);
                        WGFMU.setMeasureEvent("GATE_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, 
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        WGFMU.addVector("DRAIN_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], pMeasDrain[i]);
                        WGFMU.addVector("DRAIN_MEAS_PTN" + DutNum[index].ToString(), measWindow, pMeasDrain[i]);
                        WGFMU.setMeasureEvent("DRAIN_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, 
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if (channelSize[index] > 2)
                        {
                            WGFMU.addVector("SOURCE_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                            WGFMU.addVector("SOURCE_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                            WGFMU.setMeasureEvent("SOURCE_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, 
                                measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if(channelSize[index] > 3)
                            {
                                WGFMU.addVector("SUB_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                                WGFMU.addVector("SUB_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                                WGFMU.setMeasureEvent("SUB_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                                    measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            }                            
                        }
                    }
                }

                else // regular spot meas
                {
                    WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateSense[index]);
                    WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), measWindow, vGateSense[index]);
                    WGFMU.addVector("DRAIN_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], vDrainSense[index]); // here we will not consider any skew1 for our first measurement point
                    WGFMU.addVector("DRAIN_MEAS_PTN" + DutNum[index].ToString(), measWindow, vDrainSense[index]);
                    WGFMU.setMeasureEvent("DRAIN_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addVector("SOURCE_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0); // here we will not consider any skew1 for our first measurement point
                        WGFMU.addVector("SOURCE_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                        WGFMU.setMeasureEvent("SOURCE_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if(channelSize[index] > 3)
                        {
                            WGFMU.addVector("SUB_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0); // here we will not consider any skew1 for our first measurement point
                            WGFMU.addVector("SUB_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                            WGFMU.setMeasureEvent("SUB_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                        }                        
                    }
                }

                // create measurement pattern during stress
                // Gate
                if (acStress[index]) WGFMU.createPattern("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), vGateACLow[index]);     // AC: start at the end of the off cycle of the stress pulse
                else WGFMU.createPattern("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), vGateStress[index]);             // DC: start at the end of the DC stress voltage

                if (measAfterHigh[index] == true && acStress[index] == true)                      // if we are to measure after high then we need to add a high pulse cycle
                {
                    WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateHighTime[index], vGateStress[index]);
                }
                if (measIV[index])
                {
                    measWindowsPerSense[index] = tempNumSteps;
                    for (int i = 0; i < tempNumSteps; i++)
                    {
                        WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], ivGateStart[index] + i * ivGateStep[index]);
                        WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, ivGateStart[index] + i * ivGateStep[index]);
                    }
                }
                else if (measSpot[index])
                {
                    for (int i = 0; i < pMeasGate.Count; i++)
                    {
                        WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], pMeasGate[i]);
                        WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, pMeasGate[i]);
                    }
                }
                else // single Vg sense mesurement
                {
                    WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateSense[index]);
                    WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, vGateSense[index]);
                }

                if (invStress[index]) WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], vGateSense[index]);

                //Drain
                if (acStress[index]) WGFMU.createPattern("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), vDrainACLow[index]);   // start at the end of the off cycle of the stress pulse
                else WGFMU.createPattern("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), vDrainStress[index]);           // start at the end of the DC stress voltage

                if (channelSize[index] > 2)
                {
                    WGFMU.createPattern("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), 0);

                    if(channelSize[index] > 3)
                        WGFMU.createPattern("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), 0);
                }

                if (measAfterHigh[index] == true && acStress[index] == true)                      // if we are to measure after high then we need to add a high pulse cycle
                {
                    WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainHighTime[index], vDrainStress[index]);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);

                        if(channelSize[index] > 3)
                        {
                            WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);
                        }                       
                    }
                }
                if (measIV[index])
                {
                    measWindowsPerSense[index] = tempNumSteps;
                    for (int i = 0; i < tempNumSteps; i++)
                    {
                        WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], vDrainSense[index]);
                        WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, vDrainSense[index]);

                        if (channelSize[index] > 2)
                        {
                            WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                            WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);

                            if(channelSize[index] > 3)
                            {
                                WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                                WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                            }                            
                        }

                        if (measAfterHigh[index] == true && acStress[index] == true)
                        {
                            WGFMU.setMeasureEvent("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index] +
                                gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            WGFMU.setMeasureEvent("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index] +
                                gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if (channelSize[index] > 2)
                            {
                                WGFMU.setMeasureEvent("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index] +
                                    gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                                if(channelSize[index] > 3)
                                    WGFMU.setMeasureEvent("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index] +
                                        gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            }
                        }
                        else
                        {
                            WGFMU.setMeasureEvent("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i) + measWindow * i,
                               measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            WGFMU.setMeasureEvent("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i) + measWindow * i,
                               measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if (channelSize[index] > 2)
                            {
                                WGFMU.setMeasureEvent("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                                    measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                                if(channelSize[index] > 3)
                                    WGFMU.setMeasureEvent("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                                        measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            }
                        }
                    }
                }
                else if (measSpot[index])
                {
                    measWindowsPerSense[index] = pMeasGate.Count;
                    for (int i = 0; i < pMeasGate.Count; i++)
                    {
                        WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], pMeasDrain[i]);
                        WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, pMeasDrain[i]);

                        if (channelSize[index] > 2)
                        {
                            WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                            WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);

                            if(channelSize[index] > 3)
                            {
                                WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                                WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                            }                            
                        }

                        if (measAfterHigh[index] == true && acStress[index] == true)
                        {
                            WGFMU.setMeasureEvent("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index] +
                                gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            WGFMU.setMeasureEvent("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index] +
                                gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if (channelSize[index] > 2)
                            {
                                WGFMU.setMeasureEvent("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index] +
                                    gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                                if(channelSize[index] > 3)
                                    WGFMU.setMeasureEvent("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index] +
                                        gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            }
                        }
                        else
                        {
                            WGFMU.setMeasureEvent("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                               measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            WGFMU.setMeasureEvent("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                               measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if (channelSize[index] > 2)
                            {
                                WGFMU.setMeasureEvent("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                                    measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                                if(channelSize[index] > 3)
                                    WGFMU.setMeasureEvent("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i,
                                       measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            }
                        }
                    }
                }
                else // single measurement
                {
                    WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainSense[index]);
                    WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, vDrainSense[index]);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);

                        if(channelSize[index] > 3)
                        {
                            WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                        }                       
                    }

                    if (measAfterHigh[index] == true && acStress[index] == true)
                    {
                        WGFMU.setMeasureEvent("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index],
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                        WGFMU.setMeasureEvent("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index],
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if (channelSize[index] > 2)
                        {
                            WGFMU.setMeasureEvent("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index],
                                measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if(channelSize[index] > 3)
                                WGFMU.setMeasureEvent("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + drainTransTime[index] + drainHighTime[index],
                                    measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                        }
                    }
                    else
                    {
                        WGFMU.setMeasureEvent("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                        WGFMU.setMeasureEvent("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if (channelSize[index] > 2)
                        {
                            WGFMU.setMeasureEvent("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if(channelSize[index] > 3)
                                WGFMU.setMeasureEvent("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                        }
                    }
                }

                if (invStress[index])
                {
                    WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainACLow[index]);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);

                        if(channelSize[index] > 3)
                            WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                    }
                }

                // create measurement pattern for relaxation
                WGFMU.createPattern("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), vGateRelax[index]);     // start at the end of the off cycle of the stress pulse
                WGFMU.createPattern("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), vDrainRelax[index]);   // start at the end of the off cycle of the stress pulse

                if (channelSize[index] > 2)
                {
                    WGFMU.createPattern("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), 0);     // start at the end of the off cycle of the stress pulse

                    if(channelSize[index] > 3)
                        WGFMU.createPattern("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), 0);   // start at the end of the off cycle of the stress pulse
                }

                if (measIV[index])
                {
                    for (int i = 0; i < tempNumSteps; i++)
                    {
                        WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], ivGateStart[index] + i * ivGateStep[index]);
                        WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, ivGateStart[index] + i * ivGateStep[index]);
                        WGFMU.setMeasureEvent("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        WGFMU.addVector("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], vDrainSense[index]);
                        WGFMU.addVector("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, vDrainSense[index]);
                        WGFMU.setMeasureEvent("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if (channelSize[index] > 2)
                        {
                            WGFMU.addVector("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                            WGFMU.addVector("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                            WGFMU.setMeasureEvent("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if(channelSize[index] > 3)
                            {
                                WGFMU.addVector("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                                WGFMU.addVector("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                                WGFMU.setMeasureEvent("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            }
                        }                            
                    }
                }
                else if (measSpot[index])
                {
                    for (int i = 0; i < pMeasGate.Count; i++)
                    {
                        WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], pMeasGate[i]);
                        WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, pMeasGate[i]);
                        WGFMU.setMeasureEvent("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        WGFMU.addVector("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], pMeasDrain[i]);
                        WGFMU.addVector("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, pMeasDrain[i]);
                        WGFMU.setMeasureEvent("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if (channelSize[index] > 2)
                        {
                            WGFMU.addVector("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                            WGFMU.addVector("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                            WGFMU.setMeasureEvent("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                            if(channelSize[index] > 3)
                            {
                                WGFMU.addVector("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0);
                                WGFMU.addVector("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                                WGFMU.setMeasureEvent("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + gateTransTime[index] * (i + 1) + measWindow * i, measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                            }
                        }                            
                    }
                }
                else // IV stair
                {
                    WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateSense[index]);
                    WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, vGateSense[index]);

                    WGFMU.addVector("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainSense[index]);
                    WGFMU.addVector("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, vDrainSense[index]);
                    WGFMU.setMeasureEvent("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addVector("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                        WGFMU.setMeasureEvent("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if(channelSize[index] > 3)
                        {
                            WGFMU.addVector("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), measWindow, 0);
                            WGFMU.setMeasureEvent("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index], measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                        }
                    }                        
                }
            }
            else if (measCP[index])// Charge Pumping measurement
            {
                // Before Stress
                WGFMU.createPattern("GATE_MEAS_PTN" + DutNum[index].ToString(), 0); // start at 0 volts
                WGFMU.createPattern("DRAIN_MEAS_PTN" + DutNum[index].ToString(), 0); // start at 0 volts

                if (channelSize[index] > 2)
                {
                    WGFMU.createPattern("SOURCE_MEAS_PTN" + DutNum[index].ToString(), 0); // start at 0 volts

                    if(channelSize[index] > 3)
                        WGFMU.createPattern("SUB_MEAS_PTN" + DutNum[index].ToString(), 0); // start at 0 volts
                }

                measWindowsPerSense[index] = cpNumSteps[index];
                double delayTime = 0;
                for (int i = 0; i < cpNumSteps[index]; i++) // here we make a pattern for each freq1. that we are going to test
                {
                    // calculate the pulse parameters
                    double cpPeriod = 1 / (cpFreqStart[index] + cpFreqStep[index] * i);
                    double cpHighTime = 0.5 * cpPeriod - cpTrans[index];
                    double cpLowTime = cpHighTime;

                    int j = 0;
                    for (j = 0; j < measWindow / cpPeriod + 1; j++)
                    {
                        WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), cpTrans[index], cpHigh[index]);
                        WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), cpHighTime, cpHigh[index]);
                        WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), cpTrans[index], cpLow[index]);
                        WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), cpLowTime, cpLow[index]);
                        if (i != 0) delayTime += cpTrans[index] + cpHighTime + cpTrans[index] + cpLowTime;
                    }

                    WGFMU.addVector("DRAIN_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);
                    WGFMU.setMeasureEvent("DRAIN_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                        measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addVector("SOURCE_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);
                        WGFMU.setMeasureEvent("SOURCE_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if(channelSize[index] > 3)
                        {
                            WGFMU.addVector("SUB_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);
                            WGFMU.setMeasureEvent("SUB_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                                measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                        }                        
                    }
                }
                WGFMU.addVector("GATE_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0); // return the voltage to 0
                WGFMU.addVector("DRAIN_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0); // add in a delay on the drain to coorspond to the gate voltage change

                if (channelSize[index] > 2)
                {
                    WGFMU.addVector("SOURCE_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0); // add in a delay on the drain to coorspond to the gate voltage change

                    if(channelSize[index] > 3)
                        WGFMU.addVector("SUB_MEAS_PTN" + DutNum[index].ToString(), gateTransTime[index], 0); // add in a delay on the drain to coorspond to the gate voltage change
                }

                // During Stress
                if (acStress[index])
                {
                    WGFMU.createPattern("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), vGateACLow[index]);
                    WGFMU.createPattern("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), vDrainACLow[index]);
                }
                else
                {
                    WGFMU.createPattern("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), vGateStress[index]);
                    WGFMU.createPattern("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), vDrainStress[index]);
                }
                if (channelSize[index] > 2)
                {
                    WGFMU.createPattern("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), 0);

                    if(channelSize[index] > 3)
                        WGFMU.createPattern("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), 0);
                }
                delayTime = 0;
                for (int i = 0; i < cpNumSteps[index]; i++) // here we make a pattern for each freq[index]. that we are going to test
                {
                    // calculate the pulse parameters
                    double cpPeriod = 1 / (cpFreqStart[index] + cpFreqStep[index] * i);
                    double cpHighTime = 0.5 * cpPeriod - cpTrans[index];
                    double cpLowTime = cpHighTime;

                    int j = 0;
                    for (j = 0; j < measWindow / cpPeriod + 1; j++)
                    {
                        WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), cpTrans[index], cpHigh[index]);
                        WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), cpHighTime, cpHigh[index]);
                        WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), cpTrans[index], cpLow[index]);
                        WGFMU.addVector("GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), cpLowTime, cpLow[index]);
                        if (i != 0) delayTime += cpTrans[index] + cpHighTime + cpTrans[index] + cpLowTime;
                    }
                    if (measAfterHigh[index] == true && acStress[index] == true)
                    {
                        WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j +
                            gateTransTime[index] + gateHighTime[index], 0);
                    }
                    else WGFMU.addVector("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);
                    WGFMU.setMeasureEvent("DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                        measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addVector("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);

                        if(channelSize[index] > 3)
                            WGFMU.addVector("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);

                        WGFMU.setMeasureEvent("SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if(channelSize[index] > 3)
                            WGFMU.setMeasureEvent("SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                                measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                    }
                }

                // During Relaxation
                WGFMU.createPattern("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), vGateRelax[index]);
                WGFMU.createPattern("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), vDrainRelax[index]);

                if (channelSize[index] > 2)
                {
                    WGFMU.createPattern("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), 0);

                    if(channelSize[index] > 3)
                        WGFMU.createPattern("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), 0);
                }

                delayTime = 0;
                for (int i = 0; i < cpNumSteps[index]; i++) // here we make a pattern for each freq[index]. that we are going to test
                {
                    // calculate the pulse parameters
                    double cpPeriod = 1 / (cpFreqStart[index] + cpFreqStep[index] * i);
                    double cpHighTime = 0.5 * cpPeriod - cpTrans[index];
                    double cpLowTime = cpHighTime;
                    int j = 0;
                    for (j = 0; j < measWindow / cpPeriod + 1; j++)
                    {
                        WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), cpTrans[index], cpHigh[index]);
                        WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), cpHighTime, cpHigh[index]);
                        WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), cpTrans[index], cpLow[index]);
                        WGFMU.addVector("GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), cpLowTime, cpLow[index]);
                        if (i != 0) delayTime += cpTrans[index] + cpHighTime + cpTrans[index] + cpLowTime;
                    }

                    WGFMU.addVector("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);
                    WGFMU.setMeasureEvent("DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                        measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addVector("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);
                        WGFMU.setMeasureEvent("SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                            measPoints[index], sampleInterval[index], sampleInterval[index], outMode);

                        if(channelSize[index] > 3)
                        {
                            WGFMU.addVector("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), (cpTrans[index] * 2 + cpHighTime + cpLowTime) * j, 0);
                            WGFMU.setMeasureEvent("SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), "evt", measDelay[index] + delayTime,
                                measPoints[index], sampleInterval[index], sampleInterval[index], outMode);
                        }                       
                    }
                }
            }

            // Build the stress pattern
            if (acStress[index])
            {
                // create base pulse pattern
                if (skew[index] > 1e-9)
                {
                    WGFMU.createPattern("GATE_PULSE_PTN" + DutNum[index].ToString(), vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateHighTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateLowTime[index], vGateACLow[index]);

                    // here is half a pulse incase we need it
                    WGFMU.createPattern("GATE_PULSE_HPTN" + DutNum[index].ToString(), vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_HPTN" + DutNum[index].ToString(), gateTransTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_PULSE_HPTN" + DutNum[index].ToString(), gateHighTime[index], vGateStress[index]);

                    // create base drain pulse pattern
                    WGFMU.createPattern("DRAIN_PULSE_PTN" + DutNum[index].ToString(), vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), skew[index], vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index] - skew[index], vDrainACLow[index]);

                    // here is half a pulse incase we need it
                    WGFMU.createPattern("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), skew[index], vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), gateTransTime[index] + gateHighTime[index] - skew[index] - drainTransTime[index], vDrainStress[index]);

                    if (channelSize[index] > 2)
                    {
                        // create base source pulse pattern
                        WGFMU.createPattern("SOURCE_PULSE_PTN" + DutNum[index].ToString(), 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), skew[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index] - skew[index], 0);

                        // here is half a pulse incase we need it
                        WGFMU.createPattern("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), 0);
                        WGFMU.addVector("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), skew[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), gateTransTime[index] + gateHighTime[index] - skew[index] - drainTransTime[index], 0);

                        if(channelSize[index] > 3)
                        {
                            // create base sub pulse pattern
                            WGFMU.createPattern("SUB_PULSE_PTN" + DutNum[index].ToString(), 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), skew[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index] - skew[index], 0);

                            // here is half a pulse incase we need it
                            WGFMU.createPattern("SUB_PULSE_HPTN" + DutNum[index].ToString(), 0);
                            WGFMU.addVector("SUB_PULSE_HPTN" + DutNum[index].ToString(), skew[index], 0);
                            WGFMU.addVector("SUB_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_HPTN" + DutNum[index].ToString(), gateTransTime[index] + gateHighTime[index] - skew[index] - drainTransTime[index], 0);
                        }                        
                    }
                }
                else if (skew[index] < -1e-9)
                {
                    WGFMU.createPattern("GATE_PULSE_PTN" + DutNum[index].ToString(), vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), -skew[index], vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateHighTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateLowTime[index] + skew[index], vGateACLow[index]);

                    // here is half a pulse incase we need it
                    WGFMU.createPattern("GATE_PULSE_HPTN" + DutNum[index].ToString(), vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_HPTN" + DutNum[index].ToString(), -skew[index], vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_HPTN" + DutNum[index].ToString(), gateTransTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index] + drainHighTime[index] + skew[index] - gateTransTime[index], vGateStress[index]);

                    // create base drain pulse pattern
                    WGFMU.createPattern("DRAIN_PULSE_PTN" + DutNum[index].ToString(), vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index], vDrainACLow[index]);

                    // here is half a pulse incase we need it
                    WGFMU.createPattern("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), drainHighTime[index], vDrainStress[index]);

                    if (channelSize[index] > 2)
                    {
                        // create base source pulse pattern
                        WGFMU.createPattern("SOURCE_PULSE_PTN" + DutNum[index].ToString(), 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index], 0);

                        // here is half a pulse incase we need it
                        WGFMU.createPattern("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), 0);
                        WGFMU.addVector("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), drainHighTime[index], 0);

                        if(channelSize[index] > 3)
                        {
                            // create base sub pulse pattern
                            WGFMU.createPattern("SUB_PULSE_PTN" + DutNum[index].ToString(), vDrainACLow[index]);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index], 0);

                            // here is half a pulse incase we need it
                            WGFMU.createPattern("SUB_PULSE_PTN" + DutNum[index].ToString(), 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);
                        }                        
                    }
                }
                else
                {
                    WGFMU.createPattern("GATE_PULSE_PTN" + DutNum[index].ToString(), vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateHighTime[index], vGateStress[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateTransTime[index], vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), gateLowTime[index], vGateACLow[index]);

                    // here is half a pulse incase we need it
                    WGFMU.createPattern("GATE_PULSE_HPTN" + DutNum[index].ToString(), vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_HPTN" + DutNum[index].ToString(), gateTransTime[index], vGateACLow[index]);
                    WGFMU.addVector("GATE_PULSE_HPTN" + DutNum[index].ToString(), gateHighTime[index], vGateStress[index]);

                    // create base drain pulse pattern
                    WGFMU.createPattern("DRAIN_PULSE_PTN" + DutNum[index].ToString(), vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index], vDrainACLow[index]);

                    // here is half a pulse incase we need it
                    WGFMU.createPattern("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), vDrainACLow[index]);
                    WGFMU.addVector("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index], vDrainStress[index]);
                    WGFMU.addVector("DRAIN_PULSE_HPTN" + DutNum[index].ToString(), gateHighTime[index] + gateTransTime[index] - drainTransTime[index], vDrainStress[index]);

                    if (channelSize[index] > 2)
                    {
                        // create base source pulse pattern
                        WGFMU.createPattern("SOURCE_PULSE_PTN" + DutNum[index].ToString(), 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index], 0);

                        // here is half a pulse incase we need it
                        WGFMU.createPattern("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), 0);
                        WGFMU.addVector("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                        WGFMU.addVector("SOURCE_PULSE_HPTN" + DutNum[index].ToString(), gateHighTime[index] + gateTransTime[index] - drainTransTime[index], 0);

                        if(channelSize[index] > 3)
                        {
                            // create base sub pulse pattern
                            WGFMU.createPattern("SUB_PULSE_PTN" + DutNum[index].ToString(), 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainHighTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), drainLowTime[index], 0);

                            // here is half a pulse incase we need it
                            WGFMU.createPattern("SUB_PULSE_HPTN" + DutNum[index].ToString(), 0);
                            WGFMU.addVector("SUB_PULSE_HPTN" + DutNum[index].ToString(), drainTransTime[index], 0);
                            WGFMU.addVector("SUB_PULSE_HPTN" + DutNum[index].ToString(), gateHighTime[index] + gateTransTime[index] - drainTransTime[index], 0);
                        }                        
                    }
                }
            }
            else //DC
            {
                // we will go with 1/2 multiples of the initial sense delay
                WGFMU.createPattern("GATE_PULSE_PTN" + DutNum[index].ToString(), vGateStress[index]);
                WGFMU.addVector("GATE_PULSE_PTN" + DutNum[index].ToString(), periodRelax, vGateStress[index]);

                WGFMU.createPattern("DRAIN_PULSE_PTN" + DutNum[index].ToString(), vDrainStress[index]);
                WGFMU.addVector("DRAIN_PULSE_PTN" + DutNum[index].ToString(), periodRelax, vDrainStress[index]);

                if (channelSize[index] > 2)
                {
                    WGFMU.createPattern("SOURCE_PULSE_PTN" + DutNum[index].ToString(), 0);
                    WGFMU.addVector("SOURCE_PULSE_PTN" + DutNum[index].ToString(), periodRelax, 0);

                    if(channelSize[index] > 3)
                    {
                        WGFMU.createPattern("SUB_PULSE_PTN" + DutNum[index].ToString(), 0);
                        WGFMU.addVector("SUB_PULSE_PTN" + DutNum[index].ToString(), periodRelax, 0);
                    }                    
                }
            }

            // Build the relaxation pattern
            // we will go with 1/2 multiples of the initial sense delay
            WGFMU.createPattern("GATE_RELAX_PTN" + DutNum[index].ToString(), vGateRelax[index]);
            WGFMU.addVector("GATE_RELAX_PTN" + DutNum[index].ToString(), periodRelax, vGateRelax[index]);

            WGFMU.createPattern("DRAIN_RELAX_PTN" + DutNum[index].ToString(), vDrainRelax[index]);
            WGFMU.addVector("DRAIN_RELAX_PTN" + DutNum[index].ToString(), periodRelax, vDrainRelax[index]);

            if (channelSize[index] > 2)
            {
                WGFMU.createPattern("SOURCE_RELAX_PTN" + DutNum[index].ToString(), 0);
                WGFMU.addVector("SOURCE_RELAX_PTN" + DutNum[index].ToString(), periodRelax, 0);

                if(channelSize[index] > 3)
                {
                    WGFMU.createPattern("SUB_RELAX_PTN" + DutNum[index].ToString(), 0);
                    WGFMU.addVector("SUB_RELAX_PTN" + DutNum[index].ToString(), periodRelax, 0);
                }                
            }

            // Now all the patterns have been built string them together
            // Initial sense measurement
            WGFMU.addSequence(channel1[index], "GATE_MEAS_PTN" + DutNum[index].ToString(), 1);
            WGFMU.addSequence(channel2[index], "DRAIN_MEAS_PTN" + DutNum[index].ToString(), 1);

            //Disable
            if (channelSize[index] > 2)
            {
                WGFMU.addSequence(channel3[index], "SOURCE_MEAS_PTN" + DutNum[index].ToString(), 1);

                if(channelSize[index] > 3)
                    WGFMU.addSequence(channel4[index], "SUB_MEAS_PTN" + DutNum[index].ToString(), 1);
            }

            // Build the Stress Part
            long pulses = 0;
            for (int i = 0; i < stressTimePoint.Count; i++)
            {
                // calc the number of pulses that will fit in the next stress window
                // This calculation can overflow if the the stress period[index] is large and the period[index] of the pulse is small...may need to add multiples

                if (i == 0)
                {
                    pulses = (long)((stressTimePoint[i]) / period[index]);

                    //----------------------------------------------------
                    if (pulses == 0)
                        pulses = 1;
                    //----------------------------------------------------
                }
                else
                {
                    pulses = (long)((stressTimePoint[i] - stressTimePoint[i - 1]) / period[index]);

                    //----------------------------------------------------
                    if (pulses == 0)
                        pulses = 1;
                    //----------------------------------------------------
                }

                // build the sequence
                WGFMU.addSequence(channel1[index], "GATE_PULSE_PTN" + DutNum[index].ToString(), pulses);
                WGFMU.addSequence(channel1[index], "GATE_STRESS_MEAS_PTN" + DutNum[index].ToString(), 1);

                WGFMU.addSequence(channel2[index], "DRAIN_PULSE_PTN" + DutNum[index].ToString(), pulses);
                WGFMU.addSequence(channel2[index], "DRAIN_STRESS_MEAS_PTN" + DutNum[index].ToString(), 1);

                if (channelSize[index] > 2)
                {
                    WGFMU.addSequence(channel3[index], "SOURCE_PULSE_PTN" + DutNum[index].ToString(), pulses);
                    WGFMU.addSequence(channel3[index], "SOURCE_STRESS_MEAS_PTN" + DutNum[index].ToString(), 1);

                    if(channelSize[index] > 3)
                    {
                        WGFMU.addSequence(channel4[index], "SUB_PULSE_PTN" + DutNum[index].ToString(), pulses);
                        WGFMU.addSequence(channel4[index], "SUB_STRESS_MEAS_PTN" + DutNum[index].ToString(), 1);
                    }                    
                }      

                // now update the stress time with how many pulses we applied (will use later for outputing the data
                if (i == 0) stressTimePoint[i] = pulses * (gateTransTime[index] + gateHighTime[index] + gateLowTime[index] + gateTransTime[index]);
                else stressTimePoint[i] = stressTimePoint[i - 1] + pulses * (gateTransTime[index] + gateHighTime[index] + gateLowTime[index] + gateTransTime[index]);
            }

            // Build the Relax Part
            if (relaxTime[index] > 0)
            {
                pulses = 0;
                for (int i = 0; i < relaxTimePoint.Count; i++)
                {
                    // calc the number of pulses that will fit in the next stress window
                    if (i == 0)
                    {
                        pulses = (long)(relaxTimePoint[i] / periodRelax);

                        //----------------------------------------------------
                        if (pulses == 0)
                            pulses = 1;
                        //----------------------------------------------------
                    }
                    else
                    {
                        pulses = (long)((relaxTimePoint[i] - relaxTimePoint[i - 1]) / periodRelax);

                        //----------------------------------------------------
                        if (pulses == 0)
                            pulses = 1;
                        //----------------------------------------------------
                    }

                    // build the sequence
                    WGFMU.addSequence(channel1[index], "GATE_RELAX_PTN" + DutNum[index].ToString(), pulses);
                    WGFMU.addSequence(channel1[index], "GATE_RELAX_MEAS_PTN" + DutNum[index].ToString(), 1);

                    WGFMU.addSequence(channel2[index], "DRAIN_RELAX_PTN" + DutNum[index].ToString(), pulses);
                    WGFMU.addSequence(channel2[index], "DRAIN_RELAX_MEAS_PTN" + DutNum[index].ToString(), 1);

                    if (channelSize[index] > 2)
                    {
                        WGFMU.addSequence(channel3[index], "SOURCE_RELAX_PTN" + DutNum[index].ToString(), pulses);
                        WGFMU.addSequence(channel3[index], "SOURCE_RELAX_MEAS_PTN" + DutNum[index].ToString(), 1);

                        if(channelSize[index] > 3)
                        {
                            WGFMU.addSequence(channel4[index], "SUB_RELAX_PTN" + DutNum[index].ToString(), pulses);
                            WGFMU.addSequence(channel4[index], "SUB_RELAX_MEAS_PTN" + DutNum[index].ToString(), 1);
                        }                        
                    }

                    // now update the stress time with how many pulses we applied (will use later for outputing the data)
                    if (i == 0) relaxTimePoint[i] = pulses * periodRelax;
                    else relaxTimePoint[i] = relaxTimePoint[i - 1] + pulses * periodRelax;
                }
            }

            // Copy stressTimePoint and relaxTimePoint elements to temp_stressTimePoint and temp_relaxTimePoint
            if (index == 0)
            {
                temp_stressTimePoint1 = stressTimePoint;
                temp_relaxTimePoint1 = relaxTimePoint;
            }
            else if (index == 1)
            {
                temp_stressTimePoint2 = stressTimePoint;
                temp_relaxTimePoint2 = relaxTimePoint;
            }
            else if (index == 2)
            {
                temp_stressTimePoint3 = stressTimePoint;
                temp_relaxTimePoint3 = relaxTimePoint;
            }
            else if (index == 3)
            {
                temp_stressTimePoint4 = stressTimePoint;
                temp_relaxTimePoint4 = relaxTimePoint;
            }
            else if (index == 4)
            {
                temp_stressTimePoint5 = stressTimePoint;
                temp_relaxTimePoint5 = relaxTimePoint;
            }
        }
    }
}