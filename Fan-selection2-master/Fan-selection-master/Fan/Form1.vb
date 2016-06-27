Imports System.Text
Imports System
Imports System.Configuration
Imports System.Math
Imports System.Collections.Generic
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices

Public Structure Tmodel
    Public Tname As String              'Name
    '0-Diameter,1-Toerental,2-Dichtheid,3-Zuigmond diameter,4-Persmond lengte,5-Breedte huis,6-Lengte spiraal,7-a,8-b,9-c,10-d,11-e,
    '12-Schoeplengte,13-Aantal schoepen,14-Breedte inwendig,15-Breedte uitwendig,16-Keeldiameter,17-Inw. dia. schoepen,18-Intrede hoek,19-Uittrede hoek......
    Public Tdata() As Double
    Public Teff() As Double                 'Rendement [%]
    Public Tverm() As Double                'Vermogen[kW]
    Public TPstat() As Double               'Statische druk
    Public TPdyn() As Double                'Dynamische druk
    Public TPtot() As Double                'Totale druk
    Public TFlow() As Double                'Debiet[m3/s]
    Public werkp_opT() As Double            'rendement, P_totaal [Pa],P_statisch [Pa], as_vermogen [kW], debiet[m3/sec]
    Public Geljon() As Double               'A,B,C,E,F,G,q0_min,q0_max
    Public TFlow_scaled() As Double         'Debiet[m3/s]   (scale rules are applied)
    Public TPstat_scaled() As Double        'Statische druk (scale rules are applied)
    Public TPtot_scaled() As Double         'Totale druk    (scale rules are applied)
    Public Teff_scaled() As Double          'Rendement [%]  (scale rules are applied)
    Public Tverm_scaled() As Double         'Vermogen[kW]   (scale rules are applied)
End Structure

'Compressor stages
'1= inlet impellar
'2= outlet impellar

Public Structure Stage
    Public Typ As Integer       'Impellar type
    Public Dia0 As Integer      'Impellar diameter [m] Tschets
    Public Dia1 As Integer      'Impellar diameter [m]
    Public Rpm0 As Integer      'Impellar speed [rpm] Tschets
    Public Rpm1 As Integer      'Impellar speed [rpm]
    Public Qkg As Double        'Debiet [kg/s]
    Public Q0 As Double         'Debiet inlet [Am3/s] Tschets
    Public Q1 As Double         'Debiet inlet [Am3/s]
    Public Q2 As Double         'Debiet outlet [Am3/s]
    Public Ro0 As Double        'Density[kg/Am3] inlet flange Tschets
    Public Ro1 As Double        'Density[kg/Am3] inlet flange fan
    Public Ro2 As Double        'Density[kg/Am3] outlet flange fan
    Public Ro3 As Double        'Density[kg/Am3] outlet flange loop
    Public Ps0 As Double        'Statische druk [Pa G] Tschets
    Public Ps1 As Double        'Statische druk [Pa G] inlet flange
    Public Ps2 As Double        'Statische druk [Pa G] outlet flange fan
    Public Ps3 As Double        'Statische druk [Pa G] outlet omloop
    Public Pt0 As Double        'Totale druk [Pa G] Tschets
    Public Pt1 As Double        'Totale druk [Pa G] inlet flange
    Public Pt2 As Double        'Totale druk [Pa G] outlet flange fan
    Public Pt3 As Double        'Totale druk [Pa G] outlet omloop
    Public T1 As Double         'Temp in [c]
    Public T2 As Double         'Temp uit [c]
    Public Om_velos As Double   'Omtreksnelheid [m/s]
    Public Reynolds As Double   'Waaier OD [-]
    Public Eff As Double        'VTK gemeten Rendement [%]
    Public Ackeret As Double    'Rendement [%]
    Public Power0 As Double     'Vermogen[kW] Tschets
    Public Power As Double      'As Vermogen[kW]
    Public zuig_dia As Double   'Zuig_diameter fan
    Public uitlaat_b As Double  'Uitlaat fan breed
    Public uitlaat_h As Double  'Uitlaat fan hoog
    Public in_velos As Double   'Snelheid @ inlaat flens [m/s]
    Public uit_velos As Double  'Snelheid @ uitlaat flens [m/s]
    Public loop_loss As Double  'Druk verlies [Pa]
    Public loop_velos As Double 'Snelheid omloop [m/s]
    Public delta_pt As Double   'Drukverhoging waaier [Pa] total
    Public delta_ps As Double   'Drukverhoging waaier [Pa] static
End Structure

Public Structure PPOINT         'Polynomial regression
    Public x As Double
    Public y As Double
End Structure

Public Structure Shaft_section  'Section of the ompeller shaft
    Public dia As Double            '[m]
    Public length As Double         '[m]
    Public k_stiffness As Double    '[N.m/rad]
End Structure


Public Class Form1
    Public Tschets(31) As Tmodel            'was 31
    Public section(4) As Shaft_section      'Impeller shaft section
    Public cp_air As Double = 1.005         'Specific heat air [kJ/kg.K]
    Public cond(10) As Stage                'Process conditions
    Public PZ(12) As PPOINT                 'Raw data, Polynomial regression
    Public BZ(5, 5) As Double               'Poly Coefficients, Polynomial regression

    '-------------------------------------------------
    'The results of the polynoom calculation are stored in case 0
    'if a case is stored then the data is transferred to case 1...8

    Public case_x_conditions(30, 12) As String 'Case name, type, dia, speed for case 1..8, 
    'case10= names, case11= units
    Public ABCDE_Psta(5) As Double          'ABCD.. formula factors For the Static pressure
    Public ABCDE_Pdyn(5) As Double          'ABCD.. formula factors For the Dynamic pressure
    Public ABCDE_Ptot(5) As Double          'ABCD.. formula factors For the Total pressure
    Public ABCDE_Pow(5) As Double           'ABCD.. formula factors For the power 
    Public ABCDE_Eff(5) As Double           'ABCD.. formula factors For the Efficiency

    Public case_x_flow(50, 8) As Double     'Data storage for case 1..8
    Public case_x_Pstat(50, 8) As Double    'Data storage forcase 1..8
    Public case_x_Ptot(50) As Double        'Data storage for case 0
    Public case_x_Power(50, 8) As Double    'Data storage for case 1..8
    Public case_x_Efficiency(50) As Double  'Data storage for case 0

    Dim Torsional_point(100, 2) As Double   'For calculation on torsional frequency

    '-----"Oude benaming;Norm:;EN10027-1;Werkstof;[mm/m1/100°C];Poisson ;kg/m3;E [Gpa];Rm (20c);Rp0.2(0c);Rp0.2(20c);Rp(50c);Rp(100c);Rp(150c);Rp(200c);Rp(250c);Rp(300c);Rp(350c);Rp(400c);Equiv-ASTM;Opmerking",
    Public Shared steel() As String =
     {"16M03;EN10028-2 UNS;16M03;1.5415;1.29;0.28;7850;192;440-590;225;265;260;235;225;215;200;170;160;150;A204-GrA;--",
   "Aluminium D54S;nt.Reg.Nr: DIN1745-1;AA5083 AIMo45Mn-H116;3.3547;2.54;0.33;2700;70;305;121;150;140;135;121;67;0;0;0;0;--;--",
   "Chromanite Alloy;Robert Zapp UNS;Cr19Mn10N;1.382;1.73;0.28;7810;200;800-1050;400;500;500;445;400;360;330;310;300;295;-;--",
   "Corten - A / B;EN10155 UNS;S355J2G1W;1.8962/63;1.29;0.28;7850;192;490-630;240;355;340;255;240;226;206;166;0;0;--;Max 300c",
   "Dillimax 690T;Dill.HuttWerke;DSE690V;1.8928;1.29;0.28;7850;192;790-940;717;690;790;740;717;698;697;687;659;638;A517-GrA;--",
   "Domex 690XPD(E);EN10149-2 UNS;S700MCD(E);1.8974;1.29;0.28;7850;192;810;675;740;765;690;675;660;640;620;580;540;--;--",
   "Duplex (Avesta-2205);EN 10088-1 UfllW;X2CrNiMoN22-5-3 saisna;1.4462;1.4;0.28;7800;200;640-950;335;460;385;360;335;315;300;0;0;0;A240-S31803;Max 300c",
   "Hastelloy-C22;DIN Nr: ASTM UNS;NiCr21Mo14W 2277 B575 N06022;2.4602;1.25;0.29;9000;205;786-800;310;370;354;338;310;283;260;248;244;241;--;--",
   "Inconel- 600;DIN Nicrofer7216 ASTM SO ;NiCr15Fe Alloy 600 B168 NiCr15Fe8 Npsepo;2.4816;1.44;0.29;8400;214;550;170;240;185;180;170;165;160;155;152;150;--;--",
   "Naxtra 70;Thyssen/DIN UNS;TSTE690V;1.8928;1.29;0.28;7850;192;790-940;635;690;700;660;635;605;585;570;550;530;A517-GrA;--",
   "P265GH;EN 10028-2 UNS;P265GH ;1.0425;1.29;0.28;7850;192;410-530;205;255;234;215;205;195;175;155;140;130;A516-Gr60;--",
   "S235JRG2;EN 10025 UNS;S235JRG2 ;1.0038;1.29;0.28;7850;192;340-470;180;195;200;190;180;170;150;130;120;110;A283-GrC;--",
   "S355J2G3;EN10025 UNS;S355J2G3;1.057;1.29;0.28;7850;192;490-630;284;315;340;304;284;255;226;206;0;0;A299;Max 300c",
   "SS 304;EN10088-2;X5CrNI18-10 S30400;1.4301;1.76;0.28;7900;200;520-750;142;210;165;157;142;127;118;110;104;98;A240-304;--",
   "SS 304L;EN10088-2;X2CrNi19-11 S30403;1.4306;1.76;0.28;7900;200;520-670;132;200;155;147;132;118;108;100;94;89;A240-304L;--",
   "SS 316;EN10088-2;X5CrNiMo17-12-2 S31600;1.4401;1.76;0.28;8000;200;520-680;162;220;180;177;162;147;137;127;120;115;A240-316;--",
   "SS 316TI;EN10088-2;X6CrNiMoTi17-12-2 S31635;1.4571;1.76;0.28;8000;200;520-690;177;220;191;185;177;167;157;145;140;135;A240-316Ti;--",
   "SS 321;EN10088-2;X6CrNiTi18-10 S32100;1.4541;1.76;0.28;7900;200;500-720;167;200;184;176;167;157;147;136;130;125;A240-321;--",
   "SS 410 ;EN 10088-1 U1S;X12Cr13 (Gegloeid) 541000;1.4006;1.15;0.28;7700;216;450-650;230;250;240;235;230;225;225;220;210;195;A240-410;--",
   "SS316L;EN10088-2;X2CrNiMo17-12-2 S31603;1.4404;1.76;0.28;8000;200;520-680;152;220;170;166;152;137;127;118;113;108;A240-316L;--",
   "SuperDuplex;--;X2CrNiMoN22-5-3 saisna;1.4501;1.4;0.28;7800;200;730-930;445;550;510;480;445;405;400;395;0;0;--;--",
   "Titanium-ür 2;ASTM UNS niN;B265/348-Gr2 R50400 785(1;3.7035;0.88;0.32;4500;107;345;177;281;245;226;177;131;99;80;0;0;--;Max 280c i.v.m verbrossing",
   "Weldox700E;EN10137-2 UNS;S690QL;1.8928;1.29;0.28;7850;192;780-930;590;700;643;600;590;580;570;560;550;540;--;--",
   "WSTE/TSTE355;EN 10028-3 UNS;P355NH/NL1;1.0565/66;1.29;0.28;7850;192;470-630;284;315;340;304;284;255;226;206;186;157;A516-Gr70;--"}

    'Motoren 1500 rpm
    'Vermogen,Toerental, Frame, Lengte, Geluid Lp
    Public Shared emotor_4P() As String = {
   "4.00; 1500; 112M; 380; 56",
   "5.50; 1500; 132S; 465; 56",
   "7.50; 1500; 132M; 505; 59",
   "11.0; 1500; 160M; 645; 62",
   "15.0; 1500; 160L; 645; 62",
   "18.5; 1500; 180M; 700; 62",
   "22.0; 1500; 180L; 700; 63",
   "30.0; 1500; 200M; 774; 63",
   "37.0; 1500; 225S; 866; 66",
   "45.0; 1500; 225S; 866; 66",
   "55.0; 1500; 250S; 875; 67",
   "75.0; 1500; 280S; 1088; 68",
   "90.0; 1500; 280S; 1088; 68",
   "110; 1500; 315S; 1204; 70",
   "132; 1500; 315S; 1204; 70",
   "160; 1500; 315S; 1204; 70",
   "200; 1500; 315M; 1315; 70",
   "250; 1500; 355S; 1594; 80",
   "315; 1500; 355SM; 1646; 80",
   "355; 1500; 355SM; 1646; 80",
   "400; 1500; 355M; 1751; 80",
   "450; 1500; 355M; 1751; 80",
   "500; 1500; 355M; 1751; 80",
   "560; 1500; 400L; 1928; 85",
   "630; 1500; 400L; 1928; 85",
   "710; 1500; 400L; 1928; 85",
   "1000; 1500; Spec; 00; 00"}

    'Motoren 3000 rpm
    'Vermogen,Toerental, Frame, Lengte, Geluid Lp
    Public Shared emotor_2P() As String = {
    "4;112M;380;67",
    "5.5;132S;465;70",
    "7.7;132S;465;70",
    "11;160M;645;69",
    "15;160M;645;69",
    "18.8;160L;645;69",
    "22;180M;700;69",
    "30;200M;774;72",
    "37;200M;774;72",
    "45;225S;866;74",
    "55;250S;875;75",
    "75;280S;1088;77",
    "90;280S;1088;77",
    "110;315S;1174;78",
    "132;315S;1174;78",
    "160;315S;1174;78",
    "200;315M;1285;78",
    "250;355S;1494;83",
    "315;355SM;1546;83",
    "355;355SM;1546;83",
    "400;355M;1651;83",
    "450;355M;1651;83",
    "500;400L;1828;85",
    "560;400L;1828;85"}



    'Hz; rpm; Koppel_%,
    Public Shared EXD_VSD_torque() As String = {
     "0 ; 0; 56",
     "5 ; 149; 75",
     "10; 297; 81",
     "15; 446; 85.5",
     "20; 595; 90",
     "25; 744; 92",
     "30; 892; 94",
     "35; 1041; 96",
     "40; 1190; 98",
     "45; 1338; 100",
     "50; 1487; 90",
     "55; 1636; 83",
     "60; 1784; 76",
     "65; 1933; 69",
     "70; 2082; 63",
     "75; 2231; 57.5",
     "80; 2379; 53.5",
     "85; 2528; 49",
     "90; 2677; 46",
     "95; 2825; 43.5",
     "100; 2974; 42"}


    'Db loss voor ducting per meter
    Public Shared Duct_attenuation() As String = {
     "6x6 ;   0.30; 0.2; 0.1;  0.1;  0.1; 0.1; 0.1; 0.1",
     "12x12;  0.35; 0.2; 0.1;  0.06; 0.06; 0.06; 0.06; 0.06",
     "12x24;  0.04; 0.2; 0.1;  0.05; 0.05; 0.05; 0.05; 0.05",
     "24x24;  0.25; 0.2; 0.1;  0.03; 0.03; 0.03; 0.03; 0.03",
     "48x48;  0.15; 0.1; 0.07; 0.2;  0.2; 0.2; 0.2; 0.2",
     "72x72;  0.10; 0.1; 0.05; 0.2;  0.21; 0.2; 0.2; 0.2"}

    'ASHREA 1999, chapter 46.19, Table 21
    Public Shared TLout() As String = {
     "205;4.6;26;45;53;55;52;44;35;34;0",
     "355;4.6;24;50;60;54;36;34;31;25;0",
     "560;4.6;22;47;53;37;33;33;27;25;0",
     "815;4.6;22;51;46;26;26;24;22;38;0"}


    'Sound insulation casing
    'Beschrijving,63,125,25,500,1000,2000,4000,8000
    Public Shared insulation_casing() As String = {
    "NO insulation;0;0;0;0;0;0;0;0;0",
    "50mm Rockwool, Al-plate;0;0;0;2;7;12;12;12;12",
    "50mm Rockwool, Fe-plate;0;0;0;3;10;15;16;16;16",
    "75mm Rockwool, Al-plate;0;0;0;3;8;13;13;13;13",
    "75mm Rockwool, Fe-plate;0;0;0;4;12;19;20;20;20",
    "100mm Rockwool, Al-plate;0;0;1;5;10;15;15;15;15",
    "100mm Rockwool, Fe-plate;0;0;2;6;14;21;21;21;21",
    "125mm Rockwool, Al-plate;0;1;3;7;12;17;17;17;17",
    "125mm Rockwool, Fe-plate;0;2;4;9;16;22;22;22;22",
    "150mm Rockwool, Al-plate;0;3;7;12;17;21;21;21;21",
    "150mm Rockwool, Fe-plate;0;5;9;14;19;23;23;23;23",
    "175mm Rockwool, Al-plate;0;4;9;15;21;23;23;23;23",
    "175mm Rockwool, Fe-plate;0;7;12;17;22;24;24;24;24",
    "200mm Rockwool, Al-plate;0;5;10;17;24;25;25;25;25",
    "200mm Rockwool, Fe-plate;0;8;13;19;24;25;25;25;25",
    "50mm Rockw., Al-plate, 2.6mm Foil;2,6;0;3;7;12;17;19;19;19",
    "50mm Rockw., Fe-plate, 2.6mm Foil;2,6;3;7;12;17;20;21;21;21",
    "75mm Rockw., Al-plate, 2.6mm Foil;2,6;1;3;10;15;18;20;20;20",
    "75mm Rockw., Fe-plate, 2.6mm Foil;2,6;4;8;15;20;21;21;22;22",
    "100mm Rockw., Al-plate, 2.6mm Foil;2,6;1;4;11;16;19;21;21;21",
    "100mm Rockw., Fe-plate, 2.6mm Foil;2,6;5;9;16;21;22;22;22;22",
    "125mm Rockw., Al-plate, 2.6mm Foil;2,6;2;7;13;19;22;23;23;23",
    "125mm Rockw., Fe-plate, 2.6mm Foil;2,6;6;14;19;22;23;23;23;23",
    "150mm Rockw., Al-plate, 2.6mm Foil;2,6;4;10;15;20;23;24;24;24",
    "150mm Rockw., Fe-plate, 2.6mm Foil;2,6;8;15;21;24;25;25;25;25",
    "175mm Rockw., Al-plate, 2.6mm Foil;2,6;5;12;17;21;23;24;24;24",
    "175mm Rockw., Fe-plate, 2.6mm Foil;2,6;9;16;22;25;26;26;26;26",
    "200mm Rockw., Al-plate, 2.6mm Foil;2,6;6;14;18;22;24;25;25;25",
    "200mm Rockw., Fe-plate, 2.6mm Foil;2,6;10;17;24;25;26;27;27;27"}

    'Inlet noise dempers, Standaard absorptie coulis, dikte 200mm, TROX Nederland
    'Lengte;63;125;250;500;1000;2000;4000;8000,
    Public Shared inlet_damper() As String = {
    "NO damper;0;0;0;0;0;0;0;0",
    "L=500, Slit=60;2;5;12;23;33;32;21;16",
    "L=500, Slit=80;2;4;10;20;27;26;18;12",
    "L=500, Slit=100;2;4;9;18;24;22;16;10",
    "L=500, Slit=120;2;3;8;15;20;19;13;8",
    "L=500, Slit=140;2;3;7;13;18;16;11;7",
    "L=500, Slit=160;1;2;7;12;16;14;10;7",
    "L=500, Slit=180;1;2;6;11;15;13;8;6",
    "L=500, Slit=200;1;2;6;10;14;11;7;6",
    "L=1000, Slit=60;3;10;22;34;48;48;31;22",
    "L=1000, Slit=80;3;8;19;31;44;43;27;18",
    "L=1000, Slit=100;3;7;17;29;41;39;24;15",
    "L=1000, Slit=120;2;6;15;25;35;32;20;12",
    "L=1000, Slit=140;2;5;14;23;31;27;17;10",
    "L=1000, Slit=160;2;5;13;21;28;24;14;9",
    "L=1000, Slit=180;2;5;12;20;26;21;12;8",
    "L=1000, Slit=200;2;4;11;19;24;19;11;7",
    "L=1500, Slit=60;5;14;32;47;50;50;42;28",
    "L=1500, Slit=80;4;11;27;43;50;50;35;22",
    "L=1500, Slit=100;3;10;25;40;50;50;32;19",
    "L=1500, Slit=120;3;9;22;36;46;44;26;15",
    "L=1500, Slit=140;3;8;20;33;37;37;21;13",
    "L=1500, Slit=160;2;7;19;30;32;32;18;11",
    "L=1500, Slit=180;2;6;17;28;28;28;16;10",
    "L=1500, Slit=200;2;6;17;27;25;25;14;8",
    "L=2000, Slit=60;6;17;41;50;50;50;50;33",
    "L=2000, Slit=80;5;14;36;50;50;50;44;26",
    "L=2000, Slit=100;4;13;33;50;50;50;39;22",
    "L=2000, Slit=120;4;11;29;46;50;50;31;18",
    "L=2000, Slit=140;3;10;26;42;50;47;26;15",
    "L=2000, Slit=160;3;9;24;39;49;41;22;13",
    "L=2000, Slit=180;3;8;23;37;46;36;19;11",
    "L=2000, Slit=200;2;8;22;35;44;32;16;10",
    "L=2500, Slit=60;8;20;48;50;50;50;50;37",
    "L=2500, Slit=80;6;17;42;50;50;50;50;29",
    "L=2500, Slit=100;5;15;39;50;50;50;45;25",
    "L=2500, Slit=120;5;13;35;50;50;50;36;20",
    "L=2500, Slit=140;4;11;32;50;50;50;30;17",
    "L=2500, Slit=160;3;10;29;48;50;47;25;14",
    "L=2500, Slit=180;3;10;27;45;50;42;22;12",
    "L=2500, Slit=200;3;9;26;43;50;38;19;11",
    "L=3000, Slit=60;10;23;50;50;50;50;50;40",
    "L=3000, Slit=80;8;19;49;50;50;50;50;32",
    "L=3000, Slit=100;7;17;46;50;50;50;50;27",
    "L=3000, Slit=120;5;15;40;50;50;50;41;22",
    "L=3000, Slit=140;5;13;37;50;50;50;34;18",
    "L=3000, Slit=160;4;12;34;50;50;50;29;16",
    "L=3000, Slit=180;3;11;32;50;50;48;25;13",
    "L=3000, Slit=200;3;10;30;50;50;44;21;12"}


    Dim Std_flens_dia() As Double = {71, 80, 90, 100, 112, 125, 140, 160, 180, 200, 224, 250, 280, 315,
        355, 400, 450, 500, 560, 630, 710, 800, 900, 1000, 1120, 1250, 1400, 1600, 1800, 2000, 2100,
        2200, 2400, 2600, 2800, 3000, 3100, 3200, 3300, 99999}

    'Dim R20() As Double

    'T-model, Alle gegevens bij het hoogste rendement
    Public T_eff As Double              'Efficiency max [-]
    Public T_Ptot_Pa As Double          'Pressure totaal [Pa]
    Public T_PStat_Pa As Double         'Pressure Statisch [Pa]
    Public T_Power_opt As Double        'Power optimal point [kW]
    Public T_Toerental_sec As Double    'Toerental [/sec]
    Public T_Toerental_rpm As Double    'Toerental [rpm]
    Public T_Debiet_sec As Double       'Debiet [m3/sec]
    Public T_Debiet_hr As Double        'Debiet [m3/hr] 
    Public T_Debiet_kg_sec As Double    'Debiet [kg/sec] 
    Public T_sg_gewicht As Double       'Soortelijk gewicht [kg/m3]
    Public T_diaw_m As Double           'Diameter waaier [m]
    Public T_no_schoep As Double        'Aantal schoepen [-]
    Public T_hoek_in As Double          'Schoep intrede hoek
    Public T_hoek_uit As Double         'Schoep uittrede hoek
    Public T_omtrek_s As Double         'Waaier omtreksnelhied [m/s]
    Public T_as_kw As Double            'Opgenomen vermogen [kw]
    Public T_visco_kin As Double        'Viscositeit lucht kinamatic [m2/s]
    Public T_reynolds As Double         'Reynolds waaier [-]
    Public T_air_temp As Double         'Lucht temperatuur inlet [celcius]
    Public T_spec_labour As Double      'Specifieke arbeid [J/kg]
    Public T_Totaaldruckzahl As Double  'Kental [-]
    Public T_Staticdruckzahl As Double  'Kental [-]
    Public T_Volumezahl As Double       'Kental [-]
    Public T_laufzahl As Double         'Laufzahl kengetal [-]
    Public T_Drehzahl As Double         'Drehzahl kengetal [-]
    Public T_durchmesser_zahl As Double 'Durchesserzahl kengetal [-]

    'Gewenste gegevens, Alle gegevens bij het hoogste rendement
    Public G_eff As Double              'Efficiency max [-]
    Public G_Ptot_Pa As Double          'Pressure totaal [Pa]
    Public G_Pstat_Pa As Double         'Pressure static [Pa]
    Public G_Ptot_mBar As Double        'Pressure totaal [mBar]
    Public G_Toerental_rpm As Double    'Toerental [rpm]
    Public G_Debiet_z_act_sec As Double 'Debiet zuig [Am3/sec] (A= actual)
    Public G_Debiet_z_act_hr As Double  'Debiet zuig [Am3/hr] (A= actual)
    Public G_Debiet_z_N_hr As Double    'Debiet zuig [Nm3/hr] (normal density, 1013.25 mbar, sea level, 0 celsius)
    Public Gas_mol_weight As Double     'Gas mol gewicht [kg/mol]

    Public G_Debiet_p As Double         'Debiet pers [m3/sec]
    Public G_Debiet_kg_s As Double      'Debiet [kg/sec]
    Public G_Debiet_kg_hr As Double     'Debiet [kg/hr]
    Public G_density_act_zuig As Double 'Soortelijk gewicht [kg/Am3] (actual density)
    Public G_density_act_pers As Double 'Soortelijk gewicht [kg/Am3] (actual density)
    Public G_density_act_average As Double 'Soortelijk gewicht [kg/Am3] (actual gemiddeld)
    Public G_density_N_zuig As Double   'Soortelijk gewicht [kg/Nm3] (normal density, 1013.25 mbar, sea level, 0 celsius)
    Public G_density_N_pers As Double   'Soortelijk gewicht [kg/Nm3] (normal density, 1013.25 mbar, sea level, 0 celsius)

    Public G_Totaaldruckzahl As Double  'Kental
    Public G_omtrek_s As Double         'Omvangssnelheid waaier
    Public G_diaw_m As Double           'Diameter waaier [m]
    Public G_as_kw As Double            'Opgenomen vermogen
    Public G_visco_kin As Double        'Viscositeit lucht [m2/sec]
    Public G_reynolds As Double         'Reynolds waaier [-]
    Public G_air_temp As Double         'Lucht temperatuur in [celcius]
    Public G_temp_uit_c As Double       'Lucht temperatuur uit [c]

    'Waaier direct gekoppelde aan de motor 
    Public Direct_diaw As Double             'Diameter waaier [m] berekend
    Public Direct_diaw_m_R20 As Double       'Diameter waaier [m] in de R20 reeks
    Public Direct_Toerental_rpm As Double    'Toerental [rpm] gekozen door gebruiker
    Public Direct_omtrek_s As Double         'Waaier omtreksnelhied [m/s]
    Public Direct_as_kw As Double            'Opgenomen vermogen
    Public Direct_reynolds As Double         'Reynolds waaier [-]
    Public Direct_eff As Double              'Efficiency max [-]
    Public Direct_Debiet_z_sec As Double     'Debiet zuig [m3/sec]
    Public Direct_temp_uit_c As Double       'Lucht temperatuur uit [c].

    Dim Inertia_1, Inertia_2, Inertia_3, Inertia_4 As Double    'Torsional analyses
    Dim Springstiff_1, Springstiff_2, Springstiff_3 As Double   'Torsional analyses

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hh As Integer
        Dim words() As String

        fill_array_T_schetsen()                     'Init T-schetsen info in de array plaatsen
        Find_hi_eff()                               'Determine work points

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        ComboBox1.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions
        ComboBox2.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions
        ComboBox3.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions
        ComboBox4.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions
        ComboBox5.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions
        ComboBox6.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions
        ComboBox7.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions
        ComboBox9.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions
        ComboBox10.Items.Clear()                    'Note Combobox1 contains"startup" to prevent exceptions

        For hh = 0 To (Tschets.Length - 2)            'Fill combobox 1, 2 +5 met Fan Types
            ComboBox1.Items.Add(Tschets(hh).Tname)
            ComboBox2.Items.Add(Tschets(hh).Tname)
            ComboBox7.Items.Add(Tschets(hh).Tname)
        Next hh

        '-------Fill combobox3, Steel selection------------------
        For hh = 0 To (steel.Length - 2)            'Fill combobox3 with steel data
            words = steel(hh).Split(";")
            ComboBox3.Items.Add(words(0))
        Next hh

        Label34.Text = ChrW(963) & " 0.2 @ T bedrijf [N/mm]"

        '-------Fill combobox4, Motor selection------------------
        For hh = 0 To (emotor_4P.Length - 1)         'Fill combobox4 electric motor data
            words = emotor_4P(hh).Split(";")
            ComboBox4.Items.Add(words(0))
            ComboBox6.Items.Add(words(0))
        Next hh

        For hh = 0 To (insulation_casing.Length - 1)  'Fill combobox9 Insulation data
            words = insulation_casing(hh).Split(";")
            ComboBox9.Items.Add(words(0))
        Next hh

        For hh = 0 To (inlet_damper.Length - 1)     'Fill combobox10 inlet damper
            words = inlet_damper(hh).Split(";")
            ComboBox5.Items.Add(words(0))
            ComboBox10.Items.Add(words(0))
        Next hh

        '----------------- prevent out of bounds------------------
        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 6                 'Select T17B
        End If
        If ComboBox2.Items.Count > 0 Then
            ComboBox2.SelectedIndex = 6                 'Select T17B
        End If

        If ComboBox3.Items.Count > 0 Then
            ComboBox3.SelectedIndex = 5                 'Select Domex
        End If

        If ComboBox4.Items.Count > 0 Then
            ComboBox4.SelectedIndex = 12                'Selecteer de motor 90kW
        End If

        If ComboBox6.Items.Count > 0 Then
            ComboBox6.SelectedIndex = 12                'Selecteer de motor 90kW
        End If

        If ComboBox7.Items.Count > 0 Then
            ComboBox7.SelectedIndex = 6                 'Select T17B
        End If

        ComboBox8.SelectedIndex = 0                     'SKF a1 factor

        If ComboBox9.Items.Count > 0 Then
            ComboBox9.SelectedIndex = 5                 'Select 100m Rockwool Alu
        End If

        If ComboBox10.Items.Count > 0 Then
            ComboBox5.SelectedIndex = 0                 'Select Inlet damer
            ComboBox10.SelectedIndex = 0                'Select Inlet damer
        End If
    End Sub

    Private Sub Selectie_1()

        Dim nrq As Integer
        Dim Rel_humidity As Double
        Dim P_zuig_Pa_static As Double      'Pressure abs in [Pa] static
        Dim P_pers_Pa_static As Double      'Pressure abs in [Pa] static
        Dim P_pers_Pa_total As Double       'Pressure abs in [Pa] total
        Dim visco_temp As Double

        Dim Ttype As Int16                  'Waaier type
        Dim diam1, diam2, nn1, nn2, roo1, roo2 As Double
        Dim Power_total, flow, temp As Double

        '0-Diameter,1-Toerental,2-Dichtheid,3-Zuigmond diameter,4-Persmond lengte,5-Breedte huis,6-Lengte spiraal,7 breedte pers,8 lengte pers,9-c,10-d,11-e,
        '12-Schoeplengte,13-Aantal schoepen,14-Breedte inwendig,15-Breedte uitwendig,16-Keeldiameter,17-Inw. dia. schoepen,18-Intrede hoek,19-Uittrede hoek

        TextBox12.Text = NumericUpDown58.Value.ToString

        '--------------------- model name-----------
        Label1.Text = ""
        If RadioButton6.Checked Then Label1.Text = "2"      'Double suction

        If NumericUpDown37.Value <= 10 Then
            Label1.Text += "LD "                             'Low pressure
        Else
            If NumericUpDown37.Value > 10 And NumericUpDown37.Value <= 40 Then
                Label1.Text += "MD "                         'Medium pressure
            Else
                If NumericUpDown37.Value > 40 Then
                    Label1.Text += "HD "                     'High pressure
                End If
            End If
        End If


        nrq = ComboBox7.SelectedIndex   'Prevent out of bounds error
        If nrq >= 0 And nrq <= 30 Then
            Try
                temp = Round(NumericUpDown33.Value / 5) * 5     'Diameter waaier afronden op 5 mm
                Label1.Text += TextBox159.Text & "/" & temp.ToString & "/" & Tschets(ComboBox1.SelectedIndex).Tname

                '---------------- Debiet in [kgs/hr]------------------
                If RadioButton6.Checked Then
                    G_Debiet_kg_hr = NumericUpDown3.Value / 2
                Else
                    G_Debiet_kg_hr = NumericUpDown3.Value
                End If
                TextBox344.Text = Round(G_Debiet_kg_hr, 0).ToString
                TextBox345.Text = Round(G_Debiet_kg_hr, 0).ToString
                G_Debiet_kg_s = G_Debiet_kg_hr / 3600       'Gewenst Debiet [kg/sec]


                '------- get data from database--------------------------
                TextBox1.Text = Tschets(nrq).Tdata(0)               'Diameter waaier. 

                T_sg_gewicht = Tschets(nrq).Tdata(2)                'soortelijk gewicht lucht .
                TextBox14.Text = Round(T_sg_gewicht, 2)             'Lucht s.g.[kg/m3]
                TextBox7.Text = Tschets(nrq).Tdata(3)               'Zuigmond diameter.
                TextBox2.Text = Tschets(nrq).Tdata(4)               'Uitlaat hoogte inw. 
                TextBox3.Text = Tschets(nrq).Tdata(5)               'Uitlaat breedte inw.
                TextBox8.Text = Tschets(nrq).Tdata(6)               'Lengte spiraal

                TextBox61.Text = Tschets(nrq).Tdata(12)             'schoep lengte.
                TextBox9.Text = Tschets(nrq).Tdata(13)              'aantal schoepen.
                TextBox5.Text = Tschets(nrq).Tdata(14)              'Inw schoep breedte.
                TextBox4.Text = Tschets(nrq).Tdata(15)              'Uitw schoep breedte.
                TextBox6.Text = Tschets(nrq).Tdata(16)              'Keel diameter
                TextBox62.Text = Tschets(nrq).Tdata(17)             'inwendige schoep diameter 
                TextBox59.Text = Tschets(nrq).Tdata(18)             'schoep intrede hoek.
                TextBox60.Text = Tschets(nrq).Tdata(19)             'schoep uittrede hoek.
                TextBox342.Text = Round(Tschets(nrq).Tdata(20), 3)  'Oppervlak slakkehuis zijplaat tbv noise calc.[m2]
                TextBox53.Text = Tschets(nrq).werkp_opT(3)          'As vermogen

                T_hoek_in = Convert.ToDouble(TextBox59.Text)        'schoep intrede hoek
                T_hoek_uit = Convert.ToDouble(TextBox60.Text)       'schoep uittrede hoek
                T_eff = Tschets(nrq).werkp_opT(0) / 100             'Rendement[-]

                T_Ptot_Pa = Tschets(nrq).werkp_opT(1)               'Pressure totaal [Pa]
                T_PStat_Pa = Tschets(nrq).werkp_opT(2)              'Pressure statisch [Pa]
                T_Power_opt = Tschets(nrq).werkp_opT(3)             'AS vermogen [kW]
                T_Toerental_rpm = Tschets(nrq).Tdata(1)             'Toerental [rpm]
                T_Toerental_sec = T_Toerental_rpm / 60.0            'Toerental [/sec]

                '----------- temperaturen----------------
                T_air_temp = 20                                                 'T-schetsen proef temperatuur

                '---------- debiet----------------------
                T_Debiet_sec = Tschets(nrq).werkp_opT(4)    'Debiet [Am3/sec]
                T_Debiet_hr = Round(T_Debiet_sec * 3600, 0)                     'Debiet [Am3/hr]
                T_Debiet_kg_sec = T_Debiet_sec * T_sg_gewicht                   'Debiet [kg/s]

                '-------------- waaier-------------------------------
                T_diaw_m = Convert.ToDouble(TextBox1.Text) / 1000               'Diameter waaier [m]
                T_as_kw = Convert.ToDouble(TextBox53.Text)                      'as_vermogen [kW]
                T_omtrek_s = T_diaw_m * PI * T_Toerental_sec                    'Omtrek snelheid 

                '---------- Specifieke arbeid, pagina 25 -----------------   
                T_spec_labour = T_Ptot_Pa / T_sg_gewicht                        'Spec arbeid [J/kg]

                '------------- visco----------------
                T_visco_kin = kin_visco_air(T_air_temp)                         'Kin viscositeit [m2/s] T_schets
                visco_temp = Round(T_visco_kin * 10 ^ 6, 2)
                TextBox69.Text = visco_temp.ToString        'Visco T_schets

                '-------------------------------------------------------------------------------------------------
                '---------- specifiek toerental kengetal [-] formule 2.2 pagina 40 -------------------------------
                T_Drehzahl = T_Toerental_sec * Sqrt(T_Debiet_sec / Pow(T_spec_labour / 9.81, 0.75))


                '---------- specifiek laufzahl kengetal [-] formule 2.1 pagina 40 --------------------------------
                T_laufzahl = T_Drehzahl / 157.8


                '---------- diameter toeretal kengetal [-] formule 2.5 pagina 41 ---------------------------------
                T_durchmesser_zahl = T_diaw_m * Pow(2 * T_spec_labour / T_Debiet_sec ^ 2, 0.25) * Sqrt(PI) / 2


                '--------------- Totaldruckzahl (Zie hoofdstuk 4.2,  pagina 130 )---------------------------------
                T_Totaaldruckzahl = 2 * T_Ptot_Pa / (T_sg_gewicht * T_omtrek_s ^ 2)

                '--------------- Staicdruckzahl (Zie hoofdstuk 4.2,  pagina 130 )---------------------------------
                T_Staticdruckzahl = 2 * T_PStat_Pa / (T_sg_gewicht * T_omtrek_s ^ 2)

                '----------- Volume zahl----------------------------------------------------------------------------
                T_Volumezahl = 4 * T_Debiet_sec / (Pow(PI, 2) * Pow(T_diaw_m, 3) * T_Toerental_sec)

                '------------ Reynolds T-schets--------------------------------------------------------------------
                T_reynolds = Round(T_omtrek_s * T_diaw_m / T_visco_kin, 0)
                TextBox68.Text = Round((T_reynolds * 10 ^ -6), 2).ToString

                '-----------------Present T-model info----------------------------------------------
                TextBox10.Text = T_eff.ToString                             'Rendement
                TextBox11.Text = Round(T_Ptot_Pa, 0).ToString
                TextBox55.Text = Round(T_PStat_Pa, 0).ToString
                TextBox31.Text = T_Toerental_rpm.ToString                   '[rpm]
                TextBox13.Text = T_Debiet_sec.ToString
                TextBox56.Text = T_Debiet_hr.ToString
                TextBox30.Text = Round(T_omtrek_s, 1).ToString
                TextBox216.Text = Round(T_omtrek_s / 333, 2).ToString       'Top speed Mach number
                TextBox71.Text = T_air_temp.ToString
                TextBox17.Text = Round(T_Totaaldruckzahl, 3).ToString       'Totaldruckzahl
                TextBox18.Text = Round(T_Volumezahl, 3).ToString            'Volume zahl
                TextBox85.Text = Round(T_spec_labour, 1).ToString           'Specifieke arbeid [-]
                TextBox86.Text = Round(T_laufzahl, 3).ToString              'Laufzahl [-]
                TextBox87.Text = Round(T_Debiet_kg_sec, 2).ToString         'Debiet [kg/hr]
                TextBox88.Text = Round(T_durchmesser_zahl, 2).ToString      'Diameter kengetal [-]
                TextBox123.Text = Round(T_Drehzahl * 60, 0).ToString        'Spez Drehzahl [rpm]
                TextBox124.Text = Round(T_spec_labour, 0).ToString          'Spez. Arbeid [j/Kg]

                '--------------------------- gewenste gegevens------------------------------------------
                G_Pstat_Pa = NumericUpDown37.Value * 100    'Gewenst Pressure totaal [mbar]->[Pa]
                G_air_temp = NumericUpDown4.Value           'Gewenste arbeids temperatuur in [c]

                '------------ Gas mol Weight vochtigheid ---------------
                Gas_mol_weight = NumericUpDown8.Value / 1000        'Mol gewicht [kg/mol]

                '------------ Relatieve vochtigheid ---------------
                Rel_humidity = Convert.ToDouble(NumericUpDown5.Value)


                '----------- Zuigdruk ----------------
                P_zuig_Pa_static = NumericUpDown76.Value * 100                      '[Pa abs]
                TextBox81.Text = Round((P_zuig_Pa_static - 101325) / 100, 2)        '[mbar g]
                TextBox83.Text = Round((P_zuig_Pa_static - 101325), 2)              '[Pa g]
                TextBox91.Text = Round((P_zuig_Pa_static - 101325) / 9.80665, 1)    '[Pa g]

                '----------- Persdruk ----------------
                P_pers_Pa_static = P_zuig_Pa_static + G_Pstat_Pa   '[Pa abs]

                '---------------Density berekenen of invullen----------------
                If RadioButton3.Checked = True Then     'Density berekenen
                    NumericUpDown12.Enabled = False
                    G_density_act_zuig = calc_sg_air(P_zuig_Pa_static, G_air_temp, Rel_humidity, Gas_mol_weight)       'Actual conditions zuig
                    G_density_N_zuig = calc_sg_air(101325, 0, Rel_humidity, Gas_mol_weight)                     'Normal conditions zuig
                    NumericUpDown12.Text = Round(G_density_act_zuig, 5).ToString                                'Density zuig
                    NumericUpDown12.BackColor = Color.White         'Density invullen
                    NumericUpDown5.Visible = True                   'Relative humidity
                    GroupBox16.Visible = True                       'Molair weight
                Else
                    NumericUpDown12.Enabled = True     'Density invullen             
                    G_density_act_zuig = NumericUpDown12.Value
                    G_density_N_zuig = Round(calc_Normal_density(NumericUpDown12.Value, P_zuig_Pa_static, G_air_temp), 4) 'Normal Conditions                   'Normal conditions zuig

                    NumericUpDown12.BackColor = Color.Yellow        'Density invullen
                    NumericUpDown5.Visible = False                  'Relative humidity  
                    GroupBox16.Visible = False                      'Molair weight
                End If

                '---------------- Debiet in m3------------------
                G_Debiet_z_act_sec = G_Debiet_kg_s / G_density_act_zuig     'Gewenst Debiet [Am3/hr]
                G_Debiet_z_act_hr = G_Debiet_z_act_sec * 3600.0             'Gewenst Debiet [Am3/hr]
                G_Debiet_z_N_hr = G_Debiet_kg_s / G_density_N_zuig * 3600   'Gewenst Debiet [Nm3/hr]

                '----------- calc diameter and rpm gewenste waaier---------------
                G_omtrek_s = Pow(2 * G_Pstat_Pa / (G_density_act_zuig * T_Staticdruckzahl), 0.5)
                G_diaw_m = Pow(4 * (G_Debiet_z_act_sec) / (PI * T_Volumezahl * G_omtrek_s), 0.5)
                G_Toerental_rpm = (G_omtrek_s / (PI * G_diaw_m)) * 60

                '---------- as vermogen gewenste waaier-----------
                G_as_kw = 0.001 * G_Debiet_z_act_sec * G_Ptot_Pa / T_eff    'Go to Kw

                '---------- temperaturen, lost power is tranferred to heat -----------
                G_temp_uit_c = G_air_temp + (G_as_kw / (cp_air * G_Debiet_kg_s))

                '----------------- Actual conditions at discharge ----------------------------
                'MessageBox.Show("G_density_act_zuig= " & G_density_act_zuig.ToString & " P1= " & P_zuig_Pa_static.ToString & " P2= " & P_pers_Pa_static.ToString & " T1= " & G_air_temp.ToString & " T2= " & G_temp_uit_c.ToString)


                G_density_act_pers = calc_density(G_density_act_zuig, P_zuig_Pa_static, P_pers_Pa_static, G_air_temp, G_temp_uit_c)
                G_Debiet_p = G_Debiet_kg_s / G_density_act_pers             'Pers Debiet [Am3/hr]

                '--------- Kinmatic viscosity air[m2/s]-----------------------
                G_visco_kin = kin_visco_air(G_air_temp)                         'Kin viscositeit [m2/s]
                visco_temp = Round(G_visco_kin * 10 ^ 6, 2)

                '------------ Reynolds waaier  -------------------------------------------------------------
                G_reynolds = Round(G_omtrek_s * G_diaw_m / G_visco_kin, 0)

                '------------ Rendement Renard Waaier (Ackeret) --------------
                If CheckBox4.Checked Then
                    G_eff = 1 - 0.5 * (1 - T_eff) * Pow((1 + (T_reynolds / G_reynolds)), 0.2)
                Else
                    G_eff = T_eff
                End If

                '---------------- VTK selectie 95.2 gegevens------------------
                TextBox208.Text = Round(G_Debiet_z_act_sec, 2).ToString     'Capaciteit actual [m3/sec]
                TextBox260.Text = NumericUpDown4.Value                      'Temperatuur inlet
                TextBox261.Text = Round(calc_Normal_density(NumericUpDown12.Value, P_zuig_Pa_static, G_air_temp), 4) 'Normal Conditions

                '---------- Calc Static + total Pressure -----------------------
                Ttype = ComboBox1.SelectedIndex
                diam1 = Tschets(Ttype).Tdata(0) / 1000      'waaier diameter [m]
                diam2 = G_diaw_m
                nn1 = Tschets(Ttype).Tdata(1)               'waaier [rpm]
                nn2 = G_Toerental_rpm
                roo1 = Tschets(Ttype).Tdata(2)              'density [kg/m3]
                roo2 = NumericUpDown12.Value

                G_Pstat_Pa = Round(Scale_rule_Pressure(Tschets(Ttype).werkp_opT(2), diam1, diam2, nn1, nn2, roo1, roo2), 0)
                G_Ptot_Pa = Round(Scale_rule_Pressure(Tschets(Ttype).werkp_opT(1), diam1, diam2, nn1, nn2, roo1, roo2), 0)
                P_pers_Pa_total = P_zuig_Pa_static + G_Ptot_Pa


                '---------- presenteren-----------------------  
                TextBox16.Text = Round(G_temp_uit_c, 0).ToString                        'Temp uit
                TextBox23.Text = Round(P_pers_Pa_static / 100, 2).ToString              'Static Pers druk in mbar abs
                TextBox152.Text = Round(P_pers_Pa_total / 100, 2).ToString              'Total druk in mbar abs
                TextBox25.Text = Round(G_density_act_pers, 4).ToString
                TextBox26.Text = Round(G_omtrek_s, 0).ToString                          'Omtrek snelheid [m/s]
                TextBox217.Text = Round(G_omtrek_s / Vel_Mach(G_air_temp), 2).ToString  'Omtrek snelheid [M]
                TextBox28.Text = Round(G_Debiet_p * 3600, 0).ToString                   'Pers debiet is kleiner dan zuig debiet door drukverhoging
                TextBox27.Text = Round(G_diaw_m * 1000, 0).ToString                     'Diameter waaier [mm]
                TextBox29.Text = Round(G_Toerental_rpm, 0).ToString

                If RadioButton6.Checked Then
                    TextBox58.Text = Round(G_as_kw * 2, 1).ToString                     'Two impellers on the shaft
                Else
                    TextBox58.Text = Round(G_as_kw, 1).ToString                         'One impeller on the shaft
                End If

                TextBox20.Text = Round(G_Debiet_z_N_hr, 0).ToString                     'Debiet [Nm3/hr]  
                TextBox22.Text = Round(G_Debiet_z_act_hr, 0).ToString                   'Debiet [Am3/hr]  
                TextBox203.Text = Round(G_Ptot_Pa / 100, 1).ToString                    'Ptotal [mBar] 
                TextBox72.Text = Round((G_reynolds * 10 ^ -6), 2).ToString
                TextBox70.Text = visco_temp.ToString                                    'Visco T_schets
                TextBox74.Text = Round(G_eff * 100, 1).ToString                         'Efficiency



                '========================= 2de Bedrijfspunt===================================================================
                '=============================================================================================================

                '------------ Calcu Pstatic @ Star-Flow in chart1----------------------------------------------- 
                Dim star_flow, star_Ptot, star_Psta, star_pow, star_eff, start_dyn As Double
                Dim dia_zuig As Double
                Dim i As Integer




                flow = G_Debiet_z_act_sec
                star_flow = Round(flow * 3600, 0)
                star_Psta = (ABCDE_Psta(0) + ABCDE_Psta(1) * flow ^ 1 + ABCDE_Psta(2) * flow ^ 2 + ABCDE_Psta(3) * flow ^ 3 + ABCDE_Psta(4) * flow ^ 4 + ABCDE_Psta(5) * flow ^ 5).ToString
                star_Ptot = (ABCDE_Ptot(0) + ABCDE_Ptot(1) * flow ^ 1 + ABCDE_Ptot(2) * flow ^ 2 + ABCDE_Ptot(3) * flow ^ 3 + ABCDE_Ptot(4) * flow ^ 4 + ABCDE_Ptot(5) * flow ^ 5).ToString
                star_pow = (ABCDE_Pow(0) + ABCDE_Pow(1) * flow ^ 1 + ABCDE_Pow(2) * flow ^ 2 + ABCDE_Pow(3) * flow ^ 3 + ABCDE_Pow(4) * flow ^ 4 + ABCDE_Pow(5) * flow ^ 5).ToString
                star_eff = (ABCDE_Eff(0) + ABCDE_Eff(1) * flow ^ 1 + ABCDE_Eff(2) * flow ^ 2 + ABCDE_Eff(3) * flow ^ 3 + ABCDE_Eff(4) * flow ^ 4 + ABCDE_Eff(5) * flow ^ 5).ToString
                start_dyn = star_Ptot - star_Psta

                TextBox157.Text = Round(NumericUpDown3.Value, 0).ToString       'Debiet [kg/hr] 
                TextBox272.Text = star_flow                                     '[Am3/hr]
                TextBox271.Text = Round(star_Psta, 1)                           'Pstatic [mBar abs]
                TextBox273.Text = Round(star_Ptot, 1)                           'Ptotal [mBar abs]

                If RadioButton6.Checked Then
                    TextBox274.Text = Round(star_pow * 2, 0)                    'Two impellers on the shaft
                Else
                    TextBox274.Text = Round(star_pow, 0)                        'One impellers on the shaft
                End If


                TextBox275.Text = Round(star_eff, 0)
                TextBox75.Text = Round(start_dyn, 1)
                TextBox151.Text = Round(star_Psta + NumericUpDown76.Value, 2)     'Pstatic [mBar g]
                TextBox150.Text = Round(star_Ptot + NumericUpDown76.Value, 2)     'Ptotal [mBar g]

                cond(1).Typ = ComboBox1.SelectedIndex               '[-]        T_SCHETS
                cond(1).Q0 = Tschets(cond(1).Typ).werkp_opT(4)      '[Am3/s]    T_SCHETS
                cond(1).Pt0 = Tschets(cond(1).Typ).werkp_opT(1)     '[PaG]      T_SCHETS Pressure total  
                cond(1).Ps0 = Tschets(cond(1).Typ).werkp_opT(2)     '[PaG]      T_SCHETS pressure static 
                cond(1).Dia0 = Tschets(cond(1).Typ).Tdata(0)        '[mm]       T_SCHETS
                cond(1).Rpm0 = Tschets(cond(1).Typ).Tdata(1)        '[rpm]      T_SCHETS
                cond(1).Ro0 = Tschets(cond(1).Typ).Tdata(2)         '[kg/m3]    T_SCHETS density inlet flange 

                cond(1).T1 = G_air_temp                             '[c]
                cond(1).Dia1 = NumericUpDown33.Value                '[mm]
                cond(1).Rpm1 = NumericUpDown13.Value                '[rpm]
                cond(1).Ro1 = NumericUpDown12.Value                 'density [kg/m3] inlet flange

                cond(1).Typ = ComboBox1.SelectedIndex               '[-]        T_SCHETS
                cond(1).Q0 = Tschets(cond(1).Typ).werkp_opT(4)      '[Am3/s]    T_SCHETS
                cond(1).Pt0 = Tschets(cond(1).Typ).werkp_opT(1)     '[PaG]      T_SCHETS Pressure total  
                cond(1).Ps0 = Tschets(cond(1).Typ).werkp_opT(2)     '[PaG]      T_SCHETS pressure static 
                cond(1).Dia0 = Tschets(cond(1).Typ).Tdata(0)        '[mm]       T_SCHETS
                cond(1).Rpm0 = Tschets(cond(1).Typ).Tdata(1)        '[rpm]      T_SCHETS
                cond(1).Ro0 = Tschets(cond(1).Typ).Tdata(2)         '[kg/m3]    T_SCHETS density inlet flange 

                If RadioButton6.Checked Then
                    cond(1).Qkg = Round(NumericUpDown3.Value / 2, 0) '[kg/hr]   Double suction
                Else
                    cond(1).Qkg = NumericUpDown3.Value              '[kg/hr]
                End If
                cond(1).Pt1 = P_zuig_Pa_static                      '[Pa abs] inlet flange waaier #1           
                cond(1).Ps1 = P_zuig_Pa_static                      '[Pa abs] inlet flange waaier #1
                cond(1).Power0 = Tschets(cond(1).Typ).werkp_opT(3)  '[Am3/s] Tschets

                '---------- snelheden inlaat en uitlaat----------------
                '---Note; Voor Q1 en Q2 speelt het sg geen rol !!!!!!!!!!
                '------------------------------------------------------
                dia_zuig = Round(Tschets(cond(1).Typ).Tdata(3) * cond(1).Dia1 / cond(1).Dia0, 0)        'Zuigmond diameter.
                '--- Zoek de flensmaat in de flenzen array -----------
                For i = 1 To (Std_flens_dia.Length - 1)
                    If ((dia_zuig >= Std_flens_dia(i - 1)) And (dia_zuig < Std_flens_dia(i))) Then
                        cond(1).zuig_dia = Std_flens_dia(i)
                    End If
                Next


                Calc_stage(cond(1))                                 'Bereken de waaier #1  
                calc_loop_loss(cond(1))                             'Bereken de omloop verliezen  

                '-------------------------- Waaier #2 ----------------------
                cond(2) = cond(1)                                   'Kopieer de struct met gegevens
                cond(2).T1 = cond(1).T2                             '[c] uitlaat waaier#1 is inlaat waaier #2
                cond(2).Pt1 = cond(1).Ps3                           'Inlaat waaier #2 
                cond(2).Ps1 = cond(1).Ps3                           'Inlaat waaier #2
                cond(2).Ro1 = cond(1).Ro3                           'Ro Inlaat waaier #2

                Calc_stage(cond(2))                                 'Bereken de waaier #2  
                calc_loop_loss(cond(2))                             'Bereken de omloop verliezen  

                '-------------------------- Waaier #3 ----------------------
                cond(3) = cond(2)                                   'Kopieer de struct met gegevens
                cond(3).T1 = cond(2).T2                             '[c] uitlaat waaier #1 is inlaat waaier #2
                cond(3).Pt1 = cond(2).Ps3                           'Inlaat waaier #3 
                cond(3).Ps1 = cond(2).Ps3                           'Inlaat waaier #3
                cond(3).Ro1 = cond(2).Ro3                           'Ro Inlaat waaier #3

                Calc_stage(cond(3))                                 'Bereken de waaier #3   

                '------------ Rendement Waaier (Ackeret) --------------
                If CheckBox4.Checked Then
                    Direct_eff = cond(1).Ackeret
                Else
                    Direct_eff = cond(1).Eff
                End If

                TextBox54.Text = Round(cond(1).T2, 0).ToString                                  'Temp uit [c]
                TextBox76.Text = Round(cond(1).Om_velos, 0).ToString                            'Omtrek snelheid [m/s]
                TextBox218.Text = Round(cond(1).Om_velos / Vel_Mach(cond(1).T1), 2).ToString    'Omtrek snelheid [M]
                TextBox77.Text = Round((cond(1).Reynolds * 10 ^ -6), 2).ToString
                TextBox159.Text = Round(cond(1).zuig_dia, 0).ToString                           'Zuigmond diameter [mm]
                TextBox160.Text = Round(cond(1).uitlaat_h, 0).ToString                          'Uitlaat hoogte inw.[mm]
                TextBox161.Text = Round(cond(1).uitlaat_b, 0).ToString                          'Uitlaat breedte inw.[mm]
                TextBox57.Text = Round(cond(1).in_velos, 1).ToString                            'Inlaat snelheid [m/s]
                TextBox65.Text = Round(cond(1).uit_velos, 1).ToString                           'Uitlaat snelheid [m/s]

                cond(1).Q2 = cond(1).Qkg / cond(1).Ro2                                          'Debiet pers [Am3/hr]

                TextBox267.Text = Round(cond(1).Q2, 0).ToString                                 'Debiet outlet [Am3/hr]
                TextBox268.Text = Round(cond(1).Ro2, 4).ToString                                'Density outlet [kg/Am3]
                TextBox269.Text = Round(cond(1).Qkg / Convert.ToDouble(TextBox261.Text), 0).ToString    'Debiet zuig [Nm3/hr]

                '----------------------- present waaier #1 ------------------------
                TextBox163.Text = Round(cond(1).Pt1 / 100, 0).ToString          '[mbar] inlet flange
                TextBox165.Text = Round(cond(1).T2, 0).ToString                 '[c] outlet flange
                TextBox166.Text = Round(cond(1).Pt2 / 100, 0).ToString          '[mbar] outlet P_total
                TextBox162.Text = Round(cond(1).Ps2 / 100, 0).ToString          '[mbar] outlet P_static
                TextBox82.Text = Round(cond(1).delta_ps / 100, 0).ToString      '[mbar] dp P_static
                TextBox281.Text = Round(cond(1).Ro1, 3).ToString                '[kg/Am3] density fan in

                TextBox164.Text = Round(cond(1).Power, 0).ToString              '[kW]
                TextBox182.Text = Round(cond(1).loop_velos, 1).ToString         'snelheid [m/s]
                TextBox187.Text = Round(cond(1).Ro2, 3).ToString                '[kg/Am3] Density fan uit
                TextBox184.Text = Round(cond(1).Ro3, 3).ToString                '[kg/Am3] Density loop uit
                TextBox167.Text = Round(cond(1).loop_loss / 100, 1).ToString    '[mbar]

                '-------------------------- Waaier #2 ----------------------
                '----------- present  waaier #2 ------------------------
                TextBox169.Text = Round(cond(2).Pt1 / 100, 0).ToString          '[mbar] inlet flange
                TextBox172.Text = Round(cond(2).Pt2 / 100, 0).ToString          '[mbar] outlet P_total
                TextBox168.Text = Round(cond(2).Ps2 / 100, 0).ToString          '[mbar] outlet P_static
                TextBox171.Text = Round(cond(2).T2, 0).ToString                 '[c] outlet flange
                TextBox170.Text = Round(cond(2).Power, 0).ToString              '[kW]
                TextBox183.Text = Round(cond(2).loop_velos, 1).ToString         '[m/s]
                TextBox188.Text = Round(cond(2).Ro2, 3).ToString                '[kg/Am3] Density fan uit
                TextBox185.Text = Round(cond(2).Ro3, 3).ToString                '[kg/Am3] Density loop uit
                TextBox173.Text = Round(cond(2).loop_loss / 100, 1).ToString    '[mbar]
                TextBox128.Text = Round(cond(2).delta_ps / 100, 0).ToString     '[mbar] dp P_static
                TextBox280.Text = Round(cond(2).Ro1, 3).ToString                '[kg/Am3] density fan in

                '-------------------------- Waaier #3 (geen omloop)----------------------
                '----------- present  waaier #3 (geen omloop)------------------------
                TextBox175.Text = Round(cond(3).Pt1 / 100, 0).ToString          '[mbar] inlet flange
                TextBox178.Text = Round(cond(3).Pt2 / 100, 0).ToString          '[mbar] outlet P_total
                TextBox174.Text = Round(cond(3).Ps2 / 100, 0).ToString          '[mbar] outlet P_static
                TextBox177.Text = Round(cond(3).T2, 0).ToString                 '[c] outlet flange
                TextBox176.Text = Round(cond(3).Power, 0).ToString              '[kW]
                TextBox186.Text = Round(cond(3).Ro2, 3).ToString                '[kg/Am3] Density fan uit
                TextBox129.Text = Round(cond(3).delta_ps / 100, 0).ToString     '[mbar] dp P_static
                TextBox279.Text = Round(cond(3).Ro1, 3).ToString                '[kg/Am3] density fan in

                '-------------------------- Aantal trappen ----------------------
                Select Case True
                    Case RadioButton12.Checked              '1 trap
                        GroupBox18.Visible = False
                        GroupBox40.Visible = False
                        GroupBox46.Visible = False
                    Case RadioButton13.Checked              '2 traps
                        GroupBox18.Visible = True
                        GroupBox40.Visible = True
                        GroupBox46.Visible = True
                        Panel1.Visible = False
                        Panel2.Visible = False
                        Power_total = cond(1).Power + cond(2).Power                                 '[kW]
                        TextBox180.Text = Round((cond(2).Pt2 - cond(1).Pt1) / 100, 0).ToString      '[mbar] dP fan total
                        TextBox179.Text = Round((cond(2).Ps2 - cond(1).Ps1) / 100, 0).ToString      '[mbar] dP fan static
                        TextBox276.Text = Round((cond(2).Ps2) / 100, 0).ToString                    '[mbar] gauge fan static
                        TextBox277.Text = Round((cond(2).Pt2) / 100, 0).ToString                    '[mbar] gauge fan total
                        TextBox64.Text = Round(cond(2).T2, 0)
                        TextBox278.Text = Round((cond(2).Ro2), 3).ToString                          '[kg/Am3] density out
                    Case RadioButton14.Checked              '3 traps
                        GroupBox18.Visible = True
                        GroupBox40.Visible = True
                        GroupBox46.Visible = True
                        Panel1.Visible = True
                        Panel2.Visible = True
                        TextBox276.Text = Round((cond(3).Ps2) / 100, 0).ToString                   '[mbar] gauge fan static
                        TextBox277.Text = Round((cond(3).Pt2) / 100, 0).ToString                   '[mbar] gauge fan total
                        Power_total = cond(1).Power + cond(2).Power + cond(3).Power                '[kW]
                        TextBox180.Text = Round((cond(3).Pt2 - cond(1).Pt1) / 100, 0).ToString     '[mbar] dP fan total
                        TextBox179.Text = Round((cond(3).Ps2 - cond(1).Ps1) / 100, 0).ToString     '[mbar] dP fan static
                        TextBox64.Text = Round(cond(3).T2, 0)
                        TextBox278.Text = Round((cond(3).Ro2), 3).ToString                         '[kg/Am3] density out

                End Select
                TextBox181.Text = Round(Power_total, 0).ToString               '[kW]
            Catch ex As Exception
                MessageBox.Show(ex.Message)  ' Show the exception's message.
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, NumericUpDown19.ValueChanged, NumericUpDown17.ValueChanged, TextBox34.TextChanged, TabPage2.Enter, NumericUpDown20.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown32.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown41.ValueChanged, NumericUpDown40.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown2.ValueChanged
        Calc_stress_impeller()
    End Sub

    Private Sub Calc_stress_impeller()
        Dim T_type As Integer
        Dim maxrpm As Double
        Dim maxV As Double
        Dim sigma_allowed As Double
        Dim Waaier_dia, Waaier_dik, Waaier_as_gewicht, las_gewicht As Double   'Zonder naaf
        Dim labyrinth_gewicht As Double
        Dim Sch_breed As Double                                             'Schoep breed
        Dim Sch_dik, Sch_hoek, Sch_lengte, Sch_gewicht As Double            'Schoep_dik, Schoep_lang
        Dim schoepen_gewicht, aantal_schoep As Double                       'Totaal schoepen gewicht
        Dim Bodem_gewicht As Double
        Dim Voorplaat_dik As Double
        Dim Voorplaat_gewicht As Double
        Dim sg_ver_gewicht As Double    'sg vervangend gewicht
        Dim sigma_schoep As Double
        Dim sigma_bodemplaat As Double
        Dim V_omtrek As Double
        Dim n_actual As Double
        Dim Voorplaat_keel, inw_schoep_dia, gewicht_naaf As Double
        Dim J1, J2, J3, J4, J_naaf, J_tot, j_as As Double
        Dim dia_naaf, gewicht_as As Double
        Dim length_naaf, gewicht_pulley As Double
        Dim sg_staal As Double
        Dim n_krit, Waaier_gewicht As Double
        Dim Back_plate_steel_1, Back_plate_steel_2, Back_plate_steel_3 As Double
        Dim Front_plate_steel_1, Front_plate_steel_2, Front_plate_steel_3 As Double

        If Sch_hoek > 90 Then TextBox37.Text = "90"

        Double.TryParse(TextBox34.Text, sigma_allowed)
        sigma_allowed *= 0.7                                'Max 70% van sigma 0.2 (info Peter de Wildt)
        TextBox40.Text = Round(sigma_allowed, 0).ToString
        sigma_allowed *= 1000 ^ 2                           '[N/m2] niet [N/mm2] 

        Double.TryParse(TextBox33.Text, sg_staal)           '[kg/m3]

        Waaier_dia = NumericUpDown21.Value / 1000           '[m]
        Waaier_dik = NumericUpDown17.Value / 1000           '[m]
        Voorplaat_dik = NumericUpDown31.Value / 1000        '[m]
        '--------Selected type------------
        T_type = ComboBox1.SelectedIndex

        '--------Gewichten------------
        If T_type > -1 Then '------- schoepgewicht berekenen-----------
            Label128.Text = "Waaier type " & Tschets(T_type).Tname
            Label47.Text = "Waaier type " & Tschets(T_type).Tname
            aantal_schoep = Tschets(T_type).Tdata(13)
            Voorplaat_keel = Tschets(T_type).Tdata(16) / 1000 * (Waaier_dia / 1.0)      'Keel diam [m]
            inw_schoep_dia = Tschets(T_type).Tdata(17) / 1000 * (Waaier_dia / 1.0)      'inwendige schoep diameter [m]
            Sch_hoek = Tschets(T_type).Tdata(19)                                        'Uittrede hoek in graden
            Sch_dik = NumericUpDown20.Value / 1000 '[m]
            Sch_lengte = Tschets(T_type).Tdata(12) / 1000 * (Waaier_dia / 1.0)
            Sch_breed = Tschets(T_type).Tdata(15) / 1000 * (Waaier_dia / 1.0)           'Schoep breed uittrede [m]
            Sch_gewicht = Sch_lengte * Sch_breed * Sch_dik * sg_staal
        End If

        Bodem_gewicht = PI / 4 * Waaier_dia ^ 2 * Waaier_dik * sg_staal                                 'Bodem gewicht
        Voorplaat_gewicht = PI / 4 * (Waaier_dia ^ 2 - Voorplaat_keel ^ 2) * Voorplaat_dik * sg_staal   'Voorplaat gewicht (zuig gat verwaarloosd)
        labyrinth_gewicht = NumericUpDown32.Value                                                       'Labyrinth
        schoepen_gewicht = aantal_schoep * Sch_gewicht


        Double.TryParse(TextBox190.Text, gewicht_as)
        las_gewicht = NumericUpDown11.Value         '[kg] las toevoeg materiaal

        gewicht_pulley = NumericUpDown30.Value
        dia_naaf = NumericUpDown28.Value / 1000     '[m]
        length_naaf = NumericUpDown29.Value / 1000  '[m]

        gewicht_naaf = PI / 4 * dia_naaf ^ 2 * length_naaf * sg_staal
        TextBox93.Text = Round(gewicht_naaf, 1).ToString

        Waaier_as_gewicht = Bodem_gewicht + schoepen_gewicht + labyrinth_gewicht + Voorplaat_gewicht + gewicht_as + gewicht_naaf + gewicht_pulley + las_gewicht     'totaal gewicht
        Waaier_gewicht = Bodem_gewicht + schoepen_gewicht + labyrinth_gewicht + Voorplaat_gewicht      'Gewicht tbv N_krit


        '--------max toerental (beide zijden ingeklemd)-----------
        maxrpm = 0.32 * Sqrt(sigma_allowed * Sch_dik / (sg_staal * Waaier_dia * Sch_breed ^ 2 * Cos(Sch_hoek * PI / 180)))

        '--------max omtreksnelheid------------
        maxV = Sqrt(sigma_allowed * Sch_dik * Waaier_dia / (sg_staal * Sch_breed ^ 2 * Cos(Sch_hoek * PI / 180)))

        '--------vervangen soortelijk gewicht------------
        sg_ver_gewicht = sg_staal * (Bodem_gewicht + (Sch_gewicht * aantal_schoep)) / Bodem_gewicht

        '--------omtrek snelheid------------
        n_actual = NumericUpDown19.Value / 60.0
        V_omtrek = Waaier_dia * PI * n_actual

        '--------- spanning- bodemplaat formule (6.10) page 193------------
        sigma_bodemplaat = 0.83 * sg_ver_gewicht * V_omtrek ^ 2 / 1000 ^ 2 'Trekstekte in N/m2 niet N/mm2

        '--------Spanning schoep formule (6.1a) page 189----------
        sigma_schoep = (sg_staal / 2) * V_omtrek ^ 2 * Sch_breed ^ 2 * Cos(Sch_hoek * PI / 180) / (Sch_dik * Waaier_dia / 2)
        sigma_schoep /= 1000 ^ 2   'Trekstekte in N/m2 niet N/mm2

        'MessageBox.Show("sg=" & sg_staal.ToString &" snelh=" & V_omtrek.ToString &" breed=" & Sch_breed.ToString &" dik=" & Sch_dik.ToString &" dia=" & Waaier_dia.ToString &" sigma=" & sigma_schoep.ToString)

        '------------------ Traagheid (0.5 x m x r2)-----------------
        J1 = 0.5 * Bodem_gewicht * (0.5 * Waaier_dia) ^ 2
        J2 = 0.5 * Voorplaat_gewicht * ((0.5 * Waaier_dia) ^ 2 - (0.5 * Voorplaat_keel) ^ 2)
        J3 = 0.5 * labyrinth_gewicht * (0.5 * Voorplaat_keel) ^ 2
        J4 = 0.5 * schoepen_gewicht * (0.5 * (Waaier_dia + Voorplaat_keel) / 2) ^ 2
        J_naaf = 0.5 * gewicht_naaf * (dia_naaf / 2) ^ 2      'MassaTraagheid [kg.m2]
        Double.TryParse(TextBox190.Text, j_as)
        J_tot = J1 + J2 + J3 + J4 + J_naaf + j_as


        '------------ Eigen frequenties bodemplaat ---------------------------
        '------------ Roarks, 8 edition, pagina 793 --------------------------
        Back_plate_steel_1 = 10000 * 4.4 * (NumericUpDown17.Value / 25.4) / (NumericUpDown21.Value * 0.5 / 25.4) ^ 2
        Back_plate_steel_2 = 10000 * 20.33 * (NumericUpDown17.Value / 25.4) / (NumericUpDown21.Value * 0.5 / 25.4) ^ 2
        Back_plate_steel_3 = 10000 * 59.49 * (NumericUpDown17.Value / 25.4) / (NumericUpDown21.Value * 0.5 / 25.4) ^ 2

        TextBox209.Text = Round(Back_plate_steel_1, 1).ToString             'Bodemplaat 1ste eigenfrequentie staal [Hz]
        TextBox210.Text = Round(Back_plate_steel_2, 0).ToString             'Bodemplaat 2de  eigenfrequentie staal [Hz]
        TextBox204.Text = Round(Back_plate_steel_3, 0).ToString             'Bodemplaat 3de  eigenfrequentie staal [Hz]

        TextBox211.Text = Round(Back_plate_steel_1 * 60, 0).ToString        'Bodemplaat 1st eigenfrequentie staal [rpm]
        TextBox212.Text = Round(Back_plate_steel_2 * 60, 0).ToString        'Bodemplaat 2de eigenfrequentie staal [rpm]
        TextBox358.Text = Round(Back_plate_steel_3 * 60, 0).ToString        'Bodemplaat 3de eigenfrequentie staal [rpm]

        '------------ Eigen frequenties Voorplaat ---------------------------
        '------------ Roarks, 8 edition, pagina 793 --------------------------
        Front_plate_steel_1 = 10000 * 4.4 * (NumericUpDown31.Value / 25.4) / (NumericUpDown21.Value * 0.5 / 25.4) ^ 2
        Front_plate_steel_2 = 10000 * 20.33 * (NumericUpDown31.Value / 25.4) / (NumericUpDown21.Value * 0.5 / 25.4) ^ 2
        Front_plate_steel_3 = 10000 * 59.49 * (NumericUpDown31.Value / 25.4) / (NumericUpDown21.Value * 0.5 / 25.4) ^ 2

        TextBox373.Text = Round(Front_plate_steel_1, 1).ToString             'Frontplaat 1ste eigenfrequentie staal [Hz]
        TextBox372.Text = Round(Front_plate_steel_2, 0).ToString             'Frontplaat 2de  eigenfrequentie staal [Hz]
        TextBox360.Text = Round(Front_plate_steel_3, 0).ToString             'Frontplaat 2de  eigenfrequentie staal [Hz]

        TextBox371.Text = Round(Front_plate_steel_1 * 60, 0).ToString        'Frontplaat 1st eigenfrequentie staal [rpm]
        TextBox370.Text = Round(Front_plate_steel_2 * 60, 0).ToString        'Frontplaat 2de eigenfrequentie staal [rpm]
        TextBox359.Text = Round(Front_plate_steel_3 * 60, 0).ToString        'Frontplaat 3de eigenfrequentie staal [rpm]


        '------------------------------------ airfoil------------------------------------
        '--------------------------------------------------------------------------------
        Dim airf_hoog_uitw, airf_skin_plaat_dikte As Double
        Dim rib_plaat_dikte, rib_afstand As Double
        Dim airf_weerstandmoment, airf_buiten, airf_binnen, airf_gewicht As Double
        Dim airf_breed_inw, airf_hoog_uitw_inw, airf_force, airf_max_bend, airf_Q_load, airf_sigma_bend As Double
        Dim airf_area, airf_load, airf_tau, airf_hh As Double

        airf_hoog_uitw = NumericUpDown2.Value / 1000            '[m] uitwendige maat
        airf_skin_plaat_dikte = NumericUpDown39.Value / 1000    '[m] uitwendige plaat dikte
        rib_afstand = NumericUpDown40.Value / 1000              '[m] CL-CL ribben = segment breedte
        rib_plaat_dikte = NumericUpDown41.Value / 1000          '[m] rib plaat dikte

        airf_breed_inw = rib_afstand - rib_plaat_dikte          '[m] inwendige maat
        airf_hoog_uitw_inw = airf_hoog_uitw - 2 * airf_skin_plaat_dikte   '[m] inwendige maat

        airf_gewicht = (airf_hoog_uitw * rib_afstand - airf_breed_inw * airf_hoog_uitw_inw) * Sch_breed * sg_staal  'Gewicht [kg]
        airf_force = airf_gewicht * V_omtrek ^ 2 / (Waaier_dia / 2) * Cos(Sch_hoek * PI / 180)                      'Centrifugal Force [N]  F=m.v^2/r
        airf_Q_load = airf_force / Sch_breed                                                                        'Centrifugal load [N/m] 

        '---------- weerstandmoment--------------------------
        airf_buiten = 1 / 6 * rib_afstand * airf_hoog_uitw ^ 2              '[m3]
        airf_binnen = 1 / 6 * airf_breed_inw * airf_hoog_uitw_inw ^ 2       '[m3]
        airf_weerstandmoment = airf_buiten - airf_binnen                    '[m3]

        '-----------simple beam uniform loading-(dus opgelegd) ------
        '-------------- Mmax in het midden Mx=Q.L^2/8 ---------------
        airf_max_bend = 1 / 8 * airf_Q_load * Sch_breed ^ 2                 '[N.m]
        airf_sigma_bend = airf_max_bend / airf_weerstandmoment / 1000 ^ 2   '[N/mm2]

        '-------------- airfoilschuifspanning-------------------------------
        airf_area = rib_afstand * airf_skin_plaat_dikte * 2 * 1000 ^ 2      '[mm2]
        airf_load = airf_force / 2                                          '[N]
        airf_tau = airf_load / airf_area                                    '[N/mm2]

        '----------------------- airfoil Huber + Hencky ---------------------------------
        airf_hh = Sqrt(airf_sigma_bend ^ 2 + 3 * airf_tau ^ 2)

        TextBox67.Text = Round(airf_hh, 0).ToString                         '[N/mm2]

        '----------------------------- airfoil skin -------------------------------------
        '--------------------------------------------------------------------------------
        Dim airf_skin_weerstan, airf_skin_gewicht, airf_skin_force, airf_skin_Q_load As Double
        Dim airf_skin_max_bend, airf_skin_sigma_bend, airf_skin_area, airf_skin_load, airf_skin_tau, airf_skin_hh As Double

        airf_skin_weerstan = 1 / 6 * rib_afstand * airf_skin_plaat_dikte ^ 2                                '[m3] weerstandmoment
        airf_skin_gewicht = rib_afstand * airf_skin_plaat_dikte * Sch_breed * sg_staal                      '[kg]
        airf_skin_force = airf_skin_gewicht * V_omtrek ^ 2 / (Waaier_dia / 2) * Cos(Sch_hoek * PI / 180)    '[N]
        airf_skin_Q_load = airf_skin_force / Sch_breed

        '-----------Fixed beam uniform loading-(dus ingeklemd) ------
        '-------------- Mmax in het midden Mx=Q.L^2/12 ---------------
        airf_skin_max_bend = 1 / 12 * airf_skin_Q_load * rib_afstand ^ 2            '[N.m]
        airf_skin_sigma_bend = airf_skin_max_bend / airf_skin_weerstan / 1000 ^ 2   '[N/mm2]

        '-------------- airfoil skin schuifspanning-------------------------------
        airf_skin_area = rib_afstand * airf_skin_plaat_dikte * 1000 ^ 2             '[mm2]
        airf_skin_load = airf_skin_force / 2                                        '[N]
        airf_skin_tau = airf_skin_load / airf_skin_area                             '[N/mm2]

        '----------------------- airfoil skin Huber + Hencky ---------------------------------
        airf_skin_hh = Sqrt(airf_skin_sigma_bend ^ 2 + 3 * airf_skin_tau ^ 2)       '[N/mm2]
        TextBox73.Text = Round(airf_skin_hh, 0).ToString                            '[N/mm2]


        '--------Present data------------
        TextBox32.Text = Round(sigma_bodemplaat, 0).ToString
        TextBox36.Text = Round(Sch_breed * 1000, 1).ToString                'Breedte schoep
        TextBox37.Text = Round(Sch_hoek, 1).ToString                        'Uittrede hoek in graden
        TextBox42.Text = Round(Sch_gewicht, 1).ToString
        TextBox49.Text = Round(maxV, 0).ToString
        TextBox50.Text = Round(aantal_schoep, 0).ToString
        TextBox51.Text = Round(V_omtrek, 0).ToString
        TextBox44.Text = Round(sg_ver_gewicht, 0).ToString
        TextBox43.Text = Round(sigma_schoep, 0).ToString
        TextBox38.Text = Round(maxrpm * 60, 0).ToString
        TextBox45.Text = Round(Bodem_gewicht, 1).ToString
        TextBox94.Text = Round(Voorplaat_gewicht, 1).ToString

        TextBox374.Text = Round(Waaier_gewicht, 1).ToString

        TextBox95.Text = Round(schoepen_gewicht, 1).ToString
        TextBox192.Text = Round(Waaier_as_gewicht, 0).ToString
        TextBox96.Text = Round(Waaier_as_gewicht, 1).ToString
        TextBox103.Text = Round(Voorplaat_keel * 1000, 0).ToString
        TextBox104.Text = Round(Sch_lengte * 1000, 0).ToString

        TextBox105.Text = Round(J1, 1).ToString
        TextBox106.Text = Round(J2, 1).ToString
        TextBox107.Text = Round(J3, 1).ToString
        TextBox108.Text = Round(J4, 1).ToString
        TextBox92.Text = Round(J_naaf, 2).ToString          'Massa traagheid (0.5*M*R^2)
        TextBox109.Text = Round(J_tot, 1).ToString          'Massa traagheid Totaal
        NumericUpDown45.Value = Round(J_tot, 1).ToString

        '-------------- check airfoil stress safety-----------------------
        If airf_skin_hh > sigma_allowed / 1000 ^ 2 Then
            TextBox73.BackColor = Color.Red
        Else
            TextBox73.BackColor = Color.LightGreen
        End If

        '-------------- check box stress safety-----------------------
        If airf_hh > sigma_allowed / 1000 ^ 2 Then
            TextBox67.BackColor = Color.Red
        Else
            TextBox67.BackColor = Color.LightGreen
        End If

        '-------------- check schoep stress safety-----------------------
        If sigma_schoep > sigma_allowed / 1000 ^ 2 Then
            TextBox43.BackColor = Color.Red
        Else
            TextBox43.BackColor = Color.LightGreen
        End If

        '-------------- check rpm safety-----------------------
        If maxrpm > n_actual * 60 Then
            NumericUpDown19.BackColor = Color.Red
        Else
            NumericUpDown19.BackColor = Color.LightGreen
        End If

        '-------------- check bodemplaat stress safety-----------------------
        If sigma_bodemplaat > sigma_allowed / 1000 ^ 2 Then
            TextBox32.BackColor = Color.Red
        Else
            TextBox32.BackColor = Color.LightGreen
        End If

        '-------------- kritisch toerental---------------
        Double.TryParse(TextBox47.Text, n_krit)

        If n_krit < n_actual * 60 * 1.15 Then
            TextBox47.BackColor = Color.Red
        Else
            TextBox47.BackColor = Color.LightGreen
        End If

        '-------------- kritisch toerental bodemschijf-----------
        If Back_plate_steel_1 < n_actual * 1.05 Then
            TextBox211.BackColor = Color.Red
            TextBox209.BackColor = Color.Red
        Else
            TextBox211.BackColor = Color.LightGreen
            TextBox209.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub TabPage2_TextChanged(sender As Object, e As EventArgs) Handles TabPage2.TextChanged
        ' Calc_stress_impeller()
    End Sub
    'Find the waaier diameter in the Renard reeks
    Function find_Renard_R20(getal As Double)
        Dim x1, x2 As Double

        For hh = 0 To 100
            x1 = Renard_R20(hh)
            x2 = Renard_R20(hh + 1)

            If getal > x1 And getal < x2 Then
                Return (x1)
            End If
        Next hh

        Return (0) 'Return zero when somethings goes wrong
    End Function
    'Renard R20 reeks
    Function Renard_R20(getal As Double)
        Dim Ren As Double

        Ren = (10 ^ (getal / 20) / 10)
        Ren = Round(Ren, 2, MidpointRounding.AwayFromZero)
        Return (Ren)
    End Function
    'Calc the velocity of speed
    Function Vel_Mach(temp As Double)
        Dim Mach As Double
        Mach = 20.05 * Sqrt(temp + 273.15)
        Mach = Round(Mach, 2, MidpointRounding.AwayFromZero)
        Return (Mach)
    End Function


    Function kin_visco_air(temp As Double)
        Dim visco As Double

        '--------- Kinematic viscosity air[m2/s]
        '-----Kinematic viscosity = dynamic/density------------------
        ' Formula valid from -200 to +400 celcius------------------
        If temp > 400 Then MessageBox.Show("kin_visco_air(temp) too high (T > 400)")
        If temp < -200 Then MessageBox.Show("kin_visco_air(temp) too low (T < -200)")

        temp = temp + 273.15
        visco = 0.00009 * temp ^ 2 + 0.0351 * temp - 2.9294

        visco = visco * 10 ^ -6
        Return (visco)
    End Function

    Private Sub fill_array_T_schetsen()

        Tschets(0).Tname = "W.Bohl"
        Tschets(0).Tdata = {400, 4850, 1.2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1}
        Tschets(0).Teff = {0, 43, 75, 81, 79, 42, 0, 0, 0, 0, 0, 0}                 '[%]
        Tschets(0).Tverm = {4, 5.1, 5.5, 6.1, 6.6, 9.0, 0.0, 0, 0, 0, 0, 0}         '[kW]
        Tschets(0).TPstat = {10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10}        '[Pa]
        Tschets(0).TPtot = {8666, 8542, 8047, 7428, 6809, 3714, 0, 0, 0, 0, 0, 0}   '[Pa]
        Tschets(0).TFlow = {0, 0.255, 0.51, 0.67, 0.766, 1.021, 0, 0, 0, 0, 0, 0}   '[Am3/s]
        Tschets(0).werkp_opT = {81.0, 7428, 0, 6.144, 0.67}                         'rendement, P_totaal [Pa], P_statisch [Pa], as_vermogen [kW], debiet[m3/sec]
        Tschets(0).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}
        Tschets(0).TFlow_scaled = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}              '[Am3/s]
        Tschets(0).TPstat_scaled = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}             'Statische druk
        Tschets(0).TPtot_scaled = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}              'Totale druk
        Tschets(0).Tverm_scaled = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}              'Rendement[%]
        Tschets(0).Teff_scaled = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}               'Vermogen[kW]

        Tschets(1).Tname = "T1A"
        Tschets(1).Tdata = {1000, 1480, 1.205, 605.4, 832.4, 373.0, 6075.7, 702.7, 881.1, 1063.8, 878.9, 1295.1, 364.3, 12, 133.0, 133.0, 524.3, 605.4, 30, 30, 3.101816}
        Tschets(1).Teff = {0.00, 76.5, 79.5, 80.9, 82.52, 83.26, 83.7, 83.42, 82.3, 79.92, 76.2, 71.0}
        Tschets(1).Tverm = {6.1, 18.9, 20.3, 20.9, 21.4, 21.8, 22.2, 22.5, 22.6, 22.6, 22.5, 22.3}
        Tschets(1).TPstat = {3240.3, 3745.1, 3516.2, 3380.0, 3221.5, 3063.0, 2844.7, 2625.1, 2368.0, 2121.5, 1837.3, 1701.9}
        Tschets(1).TPtot = {3240.3, 3835.0, 3638.5, 3522.6, 3387.7, 3254.5, 3068.9, 2882.7, 2661.2, 2450.3, 2208.8, 2000}
        Tschets(1).TFlow = {0.00, 3.83, 4.47, 4.82, 5.21, 5.59, 6.05, 6.48, 6.92, 7.33, 7.79, 8.17}
        Tschets(1).werkp_opT = {83.5, 250, 0, 19.95, 5.0}
        Tschets(1).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(2).Tname = "T1B"
        Tschets(2).Tdata = {1000, 1480, 1.205, 679.0, 814.8, 370.4, 5679.0, 685.2, 850.6, 1043.2, 861.7, 1269.1, 416.0, 12, 129.6, 129.6, 596.3, 30, 30, 0, 3.101816}
        Tschets(2).Teff = {0.00, 15, 30.6, 47.0, 59.5, 69.0, 77.0, 81.0, 80.5, 77.0, 69.0, 55.0}
        Tschets(2).Tverm = {7.0, 9, 10.3, 13.6, 16.8, 19.5, 21.6, 23.1, 23.8, 23.5, 22.3, 20.5}
        Tschets(2).TPstat = {3179.6, 3240, 3296.0, 3392.6, 3455.4, 3460.0, 3334.4, 3041.7, 2580.4, 1999.7, 1156.9, 567.0}
        Tschets(2).TPtot = {3179.6, 3250, 3302.2, 3417.1, 3509.0, 3555.0, 3486.1, 3256.2, 2873.1, 2382.8, 1792.8, 1164.6}
        Tschets(2).TFlow = {0.00, 0.6, 0.95, 1.9, 2.85, 3.8, 4.75, 5.7, 6.65, 7.6, 8.55, 9.5}
        Tschets(2).werkp_opT = {99, 99, 9, 9, 9}
        Tschets(2).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(3).Tname = "T1E"
        Tschets(3).Tdata = {1000, 1480, 1.205, 617.3, 814.8, 370.4, 5679.0, 685.2, 866.7, 1045.7, 859.3, 1044.4, 385.2, 12, 129.6, 129.6, 512.3, 592.6, 30, 30, 3.101816}
        Tschets(3).Teff = {0.00, 59.5, 69.0, 72.5, 75.5, 77.2, 78.0, 78.5, 77.5, 74.0, 66.0, 50.0}
        Tschets(3).Tverm = {7.0, 16.8, 19.4, 20.7, 21.5, 22.3, 22.9, 23.2, 23.1, 22.3, 21.2, 19.4}
        Tschets(3).TPstat = {3179.6, 3455.4, 3452.3, 3389.0, 3279.2, 3121.7, 2934.4, 2654.5, 2413.4, 1792.8, 1156.9, 475.0}
        Tschets(3).TPtot = {3179.6, 3509.0, 3555.0, 3510.0, 3432.4, 3306.0, 3148.9, 2912.0, 2704.6, 2175.9, 1639.6, 1072.6}
        Tschets(3).TFlow = {0.00, 2.85, 3.8, 4.28, 4.75, 5.25, 5.7, 6.25, 6.65, 7.6, 8.55, 9.5}
        Tschets(3).werkp_opT = {78.5, 297, 0, 31.6, 6.25}
        Tschets(3).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(4).Tname = "T12A."
        Tschets(4).Tdata = {1000, 1480, 1.205, 897.2, 978.8, 489.4, 6107.7, 685.2, 913.5, 1151.7, 901.3, 1390.7, 243.1, 12, 187.6, 187.6, 758.6, 783.0, 21, 30, 3.475345}
        Tschets(4).Teff = {0.00, 69.0, 77.5, 79.64, 81.22, 82.08, 83.0, 82.0, 79.9, 76.0, 70.5, 40.0}
        Tschets(4).Tverm = {10.3, 26.4, 28.7, 29.1, 29.4, 29.3, 29.2, 28.8, 28.1, 27.4, 26.6, 22.9}
        Tschets(4).TPstat = {2325.7, 2766.8, 2491.5, 2387.2, 2261.6, 2111.9, 1938.1, 1729.1, 1497.0, 1237.7, 962.4, 614.8}
        Tschets(4).TPtot = {2325.7, 2878.2, 2670.2, 2591.3, 2494.7, 2376.0, 2235.3, 2064.2, 1871.4, 1654.0, 1422.9, 1221.3}
        Tschets(4).TFlow = {0.00, 6.58, 8.33, 8.9, 9.52, 10.13, 10.75, 11.4, 12.06, 12.72, 13.38, 14.35}
        Tschets(4).werkp_opT = {83.0, 82, 0, 3.3, 2.5}
        Tschets(4).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}


        Tschets(5).Tname = "T16B."
        Tschets(5).Tdata = {1000, 1480, 1.205, 291.3, 359.2, 184.5, 4708.7, 689.3, 666.0, 728.2, 631.6, 811.2, 469.4, 10, 32.5, 14.6, 240.8, 289.3, 45, 40, 1.603017}
        Tschets(5).Teff = {0.00, 48.0, 64.2, 66.04, 67.82, 69.06, 69.9, 69.24, 67.86, 65.26, 60.4, 20.0}
        Tschets(5).Tverm = {1.1, 2.0, 2.9, 3.1, 3.3, 3.5, 3.7, 3.9, 4.0, 4.1, 4.1, 4.0}
        Tschets(5).TPstat = {3976.8, 4249.0, 4042.1, 3934.2, 3816.3, 3686.6, 3531.1, 3286.1, 2998.5, 2656.9, 2272.5, 591.8}
        Tschets(5).TPtot = {3976.8, 4256.2, 4070.9, 3970.9, 3862.7, 3743.7, 3600.1, 3371.1, 3101.1, 2780.0, 2418.0, 833.5}
        Tschets(5).TFlow = {0.00, 0.23, 0.46, 0.52, 0.59, 0.65, 0.72, 0.8, 0.87, 0.96, 1.04, 1.34}
        Tschets(5).werkp_opT = {69.0, 1522, 0, 44.4, 1.5}
        Tschets(5).Geljon = {0.141944, 6.35969, -2229.46, 0.0001695, 0.201068, -13.27, 0.00148, 0.00671}

        Tschets(6).Tname = "T17B."
        Tschets(6).Tdata = {1000, 1480, 1.205, 738.3, 872.5, 402.7, 5704.7, 617.4, 735.6, 974.5, 837.6, 1273.8, 351.7, 12, 134.2, 134.2, 624.2, 644.3, 27, 30, 2.637289}
        Tschets(6).Teff = {0.00, 72.0, 79.45, 81.16, 82.3, 82.91, 83.0, 82.82, 81.98, 80.34, 77.32, 63.0}
        Tschets(6).Tverm = {5.8, 20.9, 23.1, 23.5, 23.9, 24.2, 24.4, 24.3, 24.1, 23.7, 23.1, 21.1}
        Tschets(6).TPstat = {2850.5, 3094.9, 2960.9, 2888.6, 2787.2, 2662.3, 2506.7, 2352.8, 2115.7, 1853.3, 1581.8, 868.7}
        Tschets(6).TPtot = {2850.5, 3198.0, 3125.4, 3072.8, 2993.9, 2892.7, 2763.7, 2634.4, 2432.4, 2207.2, 1972.8, 1348.8}
        Tschets(6).TFlow = {0.00, 4.64, 5.86, 6.21, 6.57, 6.94, 7.33, 7.67, 8.14, 8.6, 9.04, 10.02}
        Tschets(6).werkp_opT = {83.0, 138, 0, 7.4, 3.0}
        Tschets(6).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(7).Tname = "T20B"
        Tschets(7).Tdata = {1000, 1480, 1.205, 472.4, 570.1, 275.6, 5275.6, 708.7, 749.6, 859.1, 733.9, 1018.9, 433.1, 10, 126.8, 71.7, 389.0, 442.5, 29, 40, 2.173719}
        Tschets(7).Teff = {0.00, 73.5, 76.02, 78.07, 79.49, 79.95, 80.0, 78.77, 76.71, 74.07, 70.24, 66.0}
        Tschets(7).Tverm = {2.3, 10.2, 11.4, 12.5, 13.5, 14.6, 15.6, 16.4, 17.1, 17.7, 18.1, 18.3}
        Tschets(7).TPstat = {3799.1, 4453.1, 4308.6, 4200.2, 4059.4, 3888.2, 3668.3, 3397.3, 3091.9, 2764.6, 2397.8, 2053.5}
        Tschets(7).TPtot = {3799.1, 4520.3, 4401.7, 4321.2, 4214.4, 4081.3, 3906.7, 3685.7, 3435.1, 3167.5, 2869.2, 2589.8}
        Tschets(7).TFlow = {0.00, 1.68, 1.97, 2.25, 2.54, 2.84, 3.16, 3.47, 3.79, 4.1, 4.44, 4.73}
        Tschets(7).werkp_opT = {80, 628, 0, 16.8, 1.6}
        Tschets(7).Geljon = {0.15345, 1.44388, -116.84, 0.00019665, 0.22452, -3.1435, 0.01084, 0.028648}

        Tschets(8).Tname = "T21E"
        Tschets(8).Tdata = {1000, 1480, 1.205, 755.6, 673.3, 500.0, 5044.4, 622.2, 706.7, 844.4, 736.9, 1073.6, 242.2, 8, 160.0, 124.4, 626.7, 640.0, 35, 59, 2.082895}
        Tschets(8).Teff = {0.00, 58.0, 67.3, 70.26, 72.46, 73.28, 73.5, 72.82, 71.56, 70.16, 68.2, 66.0}
        Tschets(8).Tverm = {6.2, 17.0, 22.7, 24.7, 26.8, 28.6, 30.4, 31.7, 32.8, 33.4, 34.0, 34.6}
        Tschets(8).TPstat = {3199.6, 3596.4, 3329.8, 3186.0, 3002.4, 2797.8, 2554.7, 2311.6, 2048.7, 1831.7, 1624.6, 1413.8}
        Tschets(8).TPtot = {3199.6, 3636.5, 3432.3, 3327.4, 3192.6, 3043.9, 2868.5, 2691.3, 2500.6, 2344.7, 2202.5, 2054.2}
        Tschets(8).TFlow = {0.00, 2.77, 4.43, 5.21, 6.04, 6.87, 7.76, 8.54, 9.31, 9.92, 10.53, 11.09}
        Tschets(8).werkp_opT = {73.5, 232, 0, 5.9, 1.4}
        Tschets(8).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(9).Tname = "T21F"
        Tschets(9).Tdata = {1000, 1480, 1.205, 755.6, 673.3, 500.0, 5044.4, 622.2, 706.7, 844.4, 736.9, 1073.6, 242.2, 16, 160.0, 124.4, 626.7, 640.0, 35, 59, 2.082826}
        Tschets(9).Teff = {0.00, 69.0, 70.2, 71.2, 72.8, 73.86, 75.3, 75.02, 73.7, 70.66, 66.7, 64.0}
        Tschets(9).Tverm = {5.2, 27.8, 30.2, 32.2, 34.1, 35.9, 37.2, 38.8, 40.0, 41.4, 42.3, 41.3}
        Tschets(9).TPstat = {3125.2, 3844.5, 3658.5, 3571.6, 3472.4, 3343.5, 3150.0, 2884.6, 2587.0, 2244.7, 1885.0, 1432.1}
        Tschets(9).TPtot = {3125.2, 3974.2, 3818.5, 3765.4, 3706.8, 3622.4, 3486.6, 3289.3, 3066.1, 2804.5, 2525.4, 2195.1}
        Tschets(9).TFlow = {0.00, 4.99, 5.54, 6.1, 6.71, 7.32, 8.04, 8.81, 9.59, 10.37, 11.09, 12.09}
        Tschets(9).werkp_opT = {75.2, 286, 0, 7.1, 1.4}
        Tschets(9).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(10).Tname = "T22B"
        Tschets(10).Tdata = {1000, 1480, 1.205, 964.9, 905.3, 659.6, 4666.7, 666.7, 731.6, 912.3, 649.1, 1101.8, 243.9, 24, 321.1, 271.9, 800.0, 800.0, 20, 25, 2.326601}
        Tschets(10).Teff = {0.00, 48.6, 59.8, 63.6, 64.9, 69.5, 70.1, 75.2, 72.2, 71.5, 69.3, 64.8}
        Tschets(10).Tverm = {11.8, 20.6, 28.6, 32.8, 35.5, 37.6, 39.9, 41.4, 42.3, 42.5, 42.0, 40.7}
        Tschets(10).TPstat = {2213.0, 2616.0, 2645.0, 2660.0, 2623.0, 2550.0, 2477.0, 2323.0, 2154.0, 1883.0, 1634.0, 1224.0}
        Tschets(10).TPtot = {2213.0, 2638.0, 2711.0, 2755.0, 2741.0, 2704.0, 2660.0, 2550.0, 2418.0, 2191.0, 1979.0, 1649.0}
        Tschets(10).TFlow = {0.00, 3.79, 6.31, 7.57, 8.41, 9.67, 10.52, 12.2, 12.62, 13.88, 14.72, 15.98}
        Tschets(10).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(10).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(11).Tname = "T22C"
        Tschets(11).Tdata = {1000, 1480, 1.205, 964.9, 905.3, 659.6, 4666.7, 666.7, 731.6, 912.3, 649.1, 1101.8, 243.9, 12, 321.1, 271.9, 800.0, 800.0, 20, 25, 2.326601}
        Tschets(11).Teff = {0.00, 41.7, 50.9, 59.0, 63.4, 67.1, 70.3, 72.9, 73.9, 75.2, 74.8, 67.7}
        Tschets(11).Tverm = {17.6, 27.5, 33.7, 38.3, 39.5, 40.4, 41.0, 41.1, 41.1, 40.2, 38.7, 35.6}
        Tschets(11).TPstat = {2345.0, 2682.0, 2638.0, 2536.0, 2470.0, 2382.0, 2250.0, 2089.0, 1891.0, 1744.0, 1363.0, 865.0}
        Tschets(11).TPtot = {2345.0, 2726.0, 2719.0, 2689.0, 2645.0, 2580.0, 2492.0, 2374.0, 2220.0, 2052.0, 1810.0, 1363.0}
        Tschets(11).TFlow = {0.00, 4.21, 6.31, 8.41, 9.46, 10.52, 11.57, 12.62, 13.67, 14.72, 15.98, 17.67}
        Tschets(11).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(11).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(12).Tname = "T25B"
        Tschets(12).Tdata = {758, 1465, 1.205, 584.0, 650.0, 345.0, 4766.0, 600.0, 665.6, 804.2, 665.0, 990.0, 220.0, 12, 194.0, 122.0, 430.0, 460.0, 30, 31, 2.326601}
        Tschets(12).Teff = {0, 32.6, 49.5, 61.8, 71.5, 78.2, 82.2, 84.8, 85.1, 83.8, 79.4, 73.2}
        Tschets(12).Tverm = {1.8, 2.9, 3.9, 4.8, 5.7, 6.5, 7.2, 7.6, 7.9, 8.0, 7.9, 7.6}
        Tschets(12).TPstat = {1767, 1864, 1884, 1933, 1962, 1946, 1836, 1670, 1474, 1234, 932, 662}
        Tschets(12).TPtot = {1767, 1867, 1895, 1959, 2010, 2021, 1944, 1816, 1666, 1475, 1230, 1010}
        Tschets(12).TFlow = {0.00, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.4}
        Tschets(12).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(12).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(13).Tname = "T26"
        Tschets(13).Tdata = {1000, 1480, 1.205, 349.9, 438.1, 213.4, 5689.9, 625.9, 694.2, 768.1, 666.4, 885.5, 470.8, 10, 60.7, 28.9, 290.2, 331.4, 40, 40, 1.703013}
        Tschets(13).Teff = {0.00, 50.0, 68.4, 72.74, 75.38, 76.82, 77.5, 76.82, 74.8, 70.98, 65.5, 44.5}
        Tschets(13).Tverm = {1.7, 3.3, 4.6, 5.2, 5.7, 6.3, 6.8, 7.3, 7.7, 8.0, 8.2, 7.9}
        Tschets(13).TPstat = {4166.8, 4268.4, 4146.5, 4032.7, 3892.4, 3711.5, 3481.8, 3191.2, 2855.8, 2449.3, 2022.4, 1341.5}
        Tschets(13).TPtot = {4166.8, 4278.9, 4185.1, 4089.7, 3973.2, 3820.1, 3624.7, 3375.5, 3089.8, 2735.7, 2370.0, 1804.3}
        Tschets(13).TFlow = {0.00, 0.39, 0.76, 0.92, 1.09, 1.27, 1.45, 1.65, 1.86, 2.06, 2.27, 2.62}
        Tschets(13).werkp_opT = {77.5, 179, 0, 1.5, 0.5}
        Tschets(13).Geljon = {0.14663, 1.5415, -408.8, 0.00032837, 0.1665, -4.3091, 0.002516, 0.016904}

        Tschets(14).Tname = "T27"
        Tschets(14).Tdata = {1000, 1480, 1.205, 285.7, 347.8, 183.9, 4596.3, 677.0, 648.9, 701.9, 618.6, 792.5, 414.3, 16, 54.7, 23.6, 288.2, 511.8, 44, 60, 1.521398}
        Tschets(14).Teff = {0.00, 71.0, 73.18, 73.97, 74.37, 74.93, 75.01, 74.8, 74.04, 73.32, 72.43, 65.0}
        Tschets(14).Tverm = {1.3, 4.5, 4.9, 5.2, 5.4, 5.7, 6.0, 6.3, 6.5, 6.7, 6.9, 7.6}
        Tschets(14).TPstat = {3884.0, 4507.1, 4382.5, 4341.0, 4287.0, 4206.0, 4112.5, 4008.6, 3892.3, 3755.3, 3638.9, 3055.3}
        Tschets(14).TPtot = {3884.0, 4578.3, 4475.4, 4445.8, 4407.2, 4345.4, 4272.5, 4187.5, 4091.0, 3974.8, 3876.7, 3426.9}
        Tschets(14).TFlow = {0.00, 0.7, 0.8, 0.85, 0.91, 0.98, 1.05, 1.11, 1.17, 1.23, 1.28, 1.6}
        Tschets(14).werkp_opT = {74.0, 1043, 0, 18.7, 1.0}
        Tschets(14).Geljon = {0.16763, 0.12793, -479.67, -0.000030339, 0.27083, -9.9922, 0.0045166, 0.010324}

        Tschets(15).Tname = "T28"
        Tschets(15).Tdata = {1000, 1480, 1.205, 477.3, 477.3, 378.8, 4742.4, 643.9, 706.1, 792.4, 643.9, 882.6, 421.2, 8, 234.8, 151.5, 643.9, 369.7, 0, 0, 1.862612}
        Tschets(15).Teff = {0.00, 49.0, 60.0, 62.36, 64.0, 64.86, 65.0, 64.68, 64.0, 62.88, 60.91, 58.0}
        Tschets(15).Tverm = {11.6, 18.2, 24.5, 26.2, 28.2, 30.3, 32.2, 33.9, 36.0, 37.9, 40.1, 43.2}
        Tschets(15).TPstat = {4450.7, 4762.0, 4473.8, 4337.7, 4185.5, 4017.2, 3828.1, 3629.8, 3389.9, 3159.3, 2882.6, 2413.0}
        Tschets(15).TPtot = {4450.7, 4829.7, 4654.5, 4569.8, 4480.6, 4382.7, 4265.4, 4145.2, 3997.3, 3858.3, 3688.0, 3396.9}
        Tschets(15).TFlow = {0.00, 1.93, 3.16, 3.58, 4.04, 4.5, 4.92, 5.34, 5.8, 6.22, 6.68, 7.38}
        Tschets(15).werkp_opT = {65.0, 186, 0, 5.3, 1.4}
        Tschets(15).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(16).Tname = "T31A"
        Tschets(16).Tdata = {1000, 1480, 1.205, 770.4, 857.5, 422.2, 6229.6, 791.6, 878.6, 1060.7, 877.3, 1306.1, 403.0, 8, 197.9, 102.9, 567.3, 606.9, 20, 30, 3.155667}
        Tschets(16).Teff = {0.00, 53.02, 62.41, 77.14, 76.39, 80.74, 83.23, 84.42, 83.2, 79.59, 72.52, 61.87}
        Tschets(16).Tverm = {6.3, 13.0, 15.0, 16.9, 18.3, 19.3, 19.9, 20.1, 20.0, 19.6, 18.8, 17.6}
        Tschets(16).TPstat = {3139.1, 3319.7, 3373.8, 3397.4, 3320.6, 3129.5, 2865.9, 2561.2, 2201.5, 1797.2, 1336.2, 844.5}
        Tschets(16).TPtot = {3139.1, 3339.7, 3409.5, 3453.1, 3400.7, 3238.6, 3008.4, 2741.5, 2424.1, 2066.6, 1656.8, 1220.8}
        Tschets(16).TFlow = {0.00, 2.09, 2.78, 3.48, 4.18, 4.87, 5.57, 6.26, 6.96, 7.65, 8.35, 9.05}
        Tschets(16).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(16).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(17).Tname = "T31B"
        Tschets(17).Tdata = {1000, 1480, 1.205, 770.4, 857.5, 422.2, 6229.6, 791.6, 878.6, 1060.7, 877.3, 1306.1, 403.0, 8, 213.7, 118.7, 567.3, 606.9, 20, 30, 3.155667}
        Tschets(17).Teff = {0.00, 56.51, 65.14, 72.1, 78.58, 82.71, 84.89, 85.58, 85.38, 83.38, 80.38, 73.98}
        Tschets(17).Tverm = {6.8, 15.4, 17.2, 18.9, 20.4, 21.6, 22.4, 22.9, 23.0, 23.0, 22.6, 22.2}
        Tschets(17).TPstat = {3156.6, 3426.6, 3439.5, 3439.1, 3408.2, 3269.9, 3024.2, 2742.8, 2432.4, 2094.9, 1735.6, 1333.4}
        Tschets(17).TPtot = {3156.6, 3456.6, 3487.9, 3510.6, 3507.1, 3400.7, 3191.5, 2950.8, 2685.7, 2398.0, 2092.8, 1749.2}
        Tschets(17).TFlow = {0.00, 2.55, 3.25, 3.94, 4.64, 5.34, 6.03, 6.73, 7.42, 8.12, 8.81, 9.51}
        Tschets(17).werkp_opT = {87.7, 161, 0, 7.3, 3.0}
        Tschets(17).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(18).Tname = "T31C."
        Tschets(18).Tdata = {1000, 1480, 1.205, 770.4, 857.5, 455.1, 6229.6, 791.6, 878.6, 1060.7, 877.3, 1306.1, 403.0, 8, 246.7, 151.7, 567.3, 606.9, 20, 30, 3.155667}
        Tschets(18).Teff = {0.00, 62.34, 69.36, 75.3, 80.45, 84.07, 85.99, 86.61, 85.9, 83.74, 81.3, 77.49}
        Tschets(18).Tverm = {7.9, 18.7, 20.6, 22.4, 23.8, 24.9, 25.9, 26.7, 27.4, 27.9, 27.9, 27.5}
        Tschets(18).TPstat = {3231.6, 3591.0, 3607.8, 3594.7, 3528.8, 3380.7, 3174.7, 2943.9, 2674.4, 2374.9, 2055.9, 1703.4}
        Tschets(18).TPtot = {3231.6, 3632.7, 3669.3, 3679.8, 3641.4, 3524.6, 3353.7, 3161.8, 2935.1, 2682.2, 2413.7, 2115.4}
        Tschets(18).TFlow = {0.00, 3.25, 3.94, 4.64, 5.34, 6.03, 6.73, 7.42, 8.12, 8.81, 9.51, 10.21}
        Tschets(18).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(18).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(19).Tname = "T31D"
        Tschets(19).Tdata = {1000, 1480, 1.205, 770.4, 857.5, 521.1, 6229.6, 791.6, 878.6, 1060.7, 877.3, 1306.1, 403.0, 8, 278.4, 183.4, 567.3, 606.9, 20, 30, 3.155667}
        Tschets(19).Teff = {0.00, 48.55, 59.16, 67.74, 74.91, 80.69, 84.59, 86.61, 86.83, 84.66, 80.56, 74.37}
        Tschets(19).Tverm = {8.5, 17.0, 20.2, 23.0, 25.5, 27.5, 29.2, 30.6, 31.6, 32.4, 32.6, 32.3}
        Tschets(19).TPstat = {3351.9, 3597.3, 3689.8, 3735.3, 3709.3, 3622.4, 3446.5, 3213.1, 2922.2, 2570.2, 2152.1, 1692.1}
        Tschets(19).TPtot = {3351.9, 3613.5, 3721.6, 3787.9, 3787.9, 3732.1, 3592.6, 3400.7, 3156.6, 2856.6, 2495.6, 2098.0}
        Tschets(19).TFlow = {0.00, 2.32, 3.25, 4.18, 5.1, 6.03, 6.96, 7.89, 8.81, 9.74, 10.67, 11.6}
        Tschets(19).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(19).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(20).Tname = "T31E"
        Tschets(20).Tdata = {1000, 1480, 1.205, 770.4, 857.5, 521.1, 6229.6, 791.6, 878.6, 1060.7, 877.3, 1306.1, 403.0, 8, 310.0, 215.0, 567.3, 606.9, 20, 30, 3.155667}
        Tschets(20).Teff = {0.00, 55.18, 64.4, 71.52, 77.81, 82.51, 84.57, 85.39, 84.76, 82.84, 79.16, 73.67}
        Tschets(20).Tverm = {10.0, 20.2, 23.1, 25.9, 28.3, 30.6, 32.7, 34.6, 35.9, 36.9, 37.3, 37.0}
        Tschets(20).TPstat = {3435.6, 3722.1, 3781.1, 3789.6, 3752.7, 3661.8, 3485.5, 3262.2, 2974.4, 2636.0, 2241.9, 1793.7}
        Tschets(20).TPtot = {3435.6, 3749.5, 3828.0, 3861.2, 3854.2, 3798.4, 3662.3, 3484.5, 3247.3, 2964.8, 2631.7, 2249.7}
        Tschets(20).TFlow = {0.00, 3.02, 3.94, 4.87, 5.8, 6.73, 7.65, 8.58, 9.51, 10.44, 11.37, 12.29}
        Tschets(20).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(20).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(21).Tname = "T33."
        Tschets(21).Tdata = {1000, 1480, 1.205, 1013.2, 857.5, 637.2, 6229.6, 791.6, 877.3, 1060.7, 877.3, 1306.1, 411.6, 8, 314.0, 233.5, 659.6, 688.7, 10, 30, 3.155667}
        Tschets(21).Teff = {0.00, 64.0, 85.1, 86.98, 88.16, 89.0, 89.28, 88.52, 87.1, 84.2, 80.5, 74.0}
        Tschets(21).Tverm = {10.6, 24.3, 29.9, 30.4, 30.9, 31.2, 31.4, 31.0, 30.4, 29.5, 28.2, 26.2}
        Tschets(21).TPstat = {2902.2, 3304.3, 3007.1, 2888.2, 2725.6, 2578.8, 2355.0, 2073.5, 1835.7, 1521.0, 1223.8, 786.7}
        Tschets(21).TPtot = {2902.2, 3346.9, 3137.5, 3038.7, 2901.1, 2775.6, 2582.2, 2341.7, 2137.0, 1864.4, 1607.0, 1236.4}
        Tschets(21).TFlow = {0.00, 4.64, 8.12, 8.72, 9.42, 9.97, 10.72, 11.64, 12.34, 13.18, 13.92, 15.08}
        Tschets(21).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(21).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(22).Tname = "T34"
        Tschets(22).Tdata = {1000, 1480, 1.205, 1013.2, 857.5, 693.9, 6229.6, 791.6, 877.3, 1060.7, 877.3, 1306.1, 411.6, 8, 370.7, 290.2, 659.6, 688.7, 10, 30, 3.155667}
        Tschets(22).Teff = {0.00, 69.0, 83.2, 85.2, 87.0, 88.22, 88.9, 88.3, 87.12, 85.1, 82.6, 52.0}
        Tschets(22).Tverm = {12.2, 28.6, 33.8, 34.7, 35.2, 35.7, 35.9, 35.9, 35.4, 34.7, 33.9, 26.6}
        Tschets(22).TPstat = {2954.7, 3339.3, 2874.2, 2788.6, 2694.8, 2535.7, 2372.5, 2173.2, 1963.4, 1739.6, 1503.6, 760.5}
        Tschets(22).TPtot = {2954.7, 3400.0, 3017.8, 2954.5, 2886.3, 2754.1, 2620.8, 2452.6, 2273.6, 2084.5, 1882.7, 1334.8}
        Tschets(22).TFlow = {0.00, 6.03, 9.28, 9.97, 10.72, 11.46, 12.2, 12.94, 13.64, 14.38, 15.08, 18.56}
        Tschets(22).werkp_opT = {88.5, 157, 0, 11.8, 5.0}
        Tschets(22).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(23).Tname = "T35A"
        Tschets(23).Tdata = {1000, 1480, 1.205, 233.3, 333.3, 175.0, 5016.7, 792.0, 736.7, 760.0, 650.0, 816.7, 433.3, 8, 108.3, 33.3, 375.0, 133.3, 0, 0, 1.87903}
        Tschets(23).Teff = {0.00, 41.0, 53.0, 54.89, 56.02, 56.56, 56.9, 56.44, 55.9, 54.8, 52.56, 36.5}
        Tschets(23).Tverm = {3.4, 5.0, 6.7, 7.2, 7.9, 8.4, 9.1, 9.7, 10.4, 11.0, 11.8, 14.6}
        Tschets(23).TPstat = {4325.0, 4499.4, 4206.4, 4080.9, 3913.5, 3744.6, 3536.8, 3313.5, 3048.5, 2787.6, 2419.2, 2232.3}
        Tschets(23).TPtot = {4325.0, 4530.2, 4322.7, 4232.7, 4114.3, 3996.4, 3850.6, 3696.2, 3513.5, 3335.8, 3088.9, 2380.7}
        Tschets(23).TFlow = {0.00, 0.42, 0.82, 0.94, 1.08, 1.2, 1.34, 1.48, 1.64, 1.78, 1.96, 2.57}
        Tschets(23).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(23).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(24).Tname = "T35B"
        Tschets(24).Tdata = {1000, 1480, 1.205, 233.3, 333.3, 175.0, 5016.7, 792.0, 736.7, 760.0, 650.0, 816.7, 433.3, 8, 108.3, 50.0, 375.0, 133.3, 0, 0, 1.87903}
        Tschets(24).Teff = {0.00, 40.0, 50.0, 55.06, 57.72, 59.1, 59.6, 56.6, 52.0, 45.0, 35.0, 31.0}
        Tschets(24).Tverm = {3.4, 5.6, 6.6, 7.5, 8.5, 9.5, 10.5, 12.5, 14.0, 15.7, 17.5, 18.2}
        Tschets(24).TPstat = {4429.7, 4729.6, 4604.1, 4459.0, 4234.3, 3981.8, 3683.3, 2949.4, 2346.7, 1576.5, 558.1, 0.0}
        Tschets(24).TPtot = {4429.7, 4767.6, 4689.5, 4596.0, 4444.0, 4279.5, 4084.2, 3603.2, 3221.4, 2725.0, 2041.0, 1522.3}
        Tschets(24).TFlow = {0.00, 0.47, 0.7, 0.89, 1.1, 1.31, 1.52, 1.94, 2.24, 2.57, 2.92, 3.04}
        Tschets(24).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(24).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(25).Tname = "T35D"  'CHECH TORERENTAL
        Tschets(25).Tdata = {1000, 1480, 1.205, 303.3, 333.3, 175.0, 5016.7, 1208.3, 736.7, 760.0, 650.0, 816.7, 433.3, 8, 108.3, 50.0, 375.0, 133.3, 0, 0, 2.102013}
        Tschets(25).Teff = {0.00, 30.0, 45.5, 47.98, 50.7, 51.92, 52.1, 51.68, 50.14, 47.48, 44.0, 30.0}
        Tschets(25).Tverm = {5.0, 7.1, 9.3, 9.9, 10.9, 11.7, 12.8, 14.2, 15.6, 17.2, 18.7, 22.4}
        Tschets(25).TPstat = {4813.3, 4743.6, 4429.7, 4302.7, 4107.4, 3919.0, 3655.4, 3278.7, 2846.2, 2323.2, 1744.0, 154.8}
        Tschets(25).TPtot = {4813.3, 4781.5, 4581.5, 4503.5, 4384.1, 4272.2, 4120.4, 3901.4, 3649.5, 3339.7, 2999.1, 2016.0}
        Tschets(25).TFlow = {0.00, 0.47, 0.94, 1.08, 1.26, 1.43, 1.64, 1.89, 2.15, 2.43, 2.69, 3.27}
        Tschets(25).werkp_opT = {53.0, 600, 0, 10.5, 0.7}
        Tschets(25).Geljon = {0, 0, 0, 0, 0, 0, 0, 0}

        Tschets(26).Tname = "T36"
        Tschets(26).Tdata = {1000, 1480, 1.205, 523.7, 572.4, 328.9, 5657.9, 710.5, 750.0, 859.2, 733.6, 1019.7, 425.0, 10, 156.6, 62.5, 442.1, 464.5, 29, 40, 2.175123}
        Tschets(26).Teff = {0.00, 71.5, 80.5, 83.5, 86.52, 87.58, 88.0, 87.68, 86.2, 83.4, 79.0, 57.0}
        Tschets(26).Tverm = {3.0, 9.6, 12.0, 13.1, 14.1, 15.0, 15.9, 16.8, 17.4, 17.8, 18.6, 17.6}
        Tschets(26).TPstat = {3930.4, 4269.4, 4121.7, 4034.8, 3883.5, 3688.7, 3471.3, 3177.4, 2805.4, 2469.6, 2052.2, 1443.5}
        Tschets(26).TPtot = {3930.4, 4312.8, 4210.0, 4151.5, 4034.9, 3879.4, 3705.7, 3469.7, 3206.9, 2896.6, 2560.4, 2237.6}
        Tschets(26).TFlow = {0.00, 1.61, 2.3, 2.65, 3.01, 3.38, 3.75, 4.19, 4.63, 5.06, 5.52, 5.9}
        Tschets(26).werkp_opT = {87.9, 221, 0, 5.1, 1.5}
        Tschets(26).Geljon = {0.14499, 1.2327, -79.528, 0.00060039, 0.16925, -1.9817, 0.01039, 0.03562}

        Tschets(27).Tname = "T36A."
        Tschets(27).Tdata = {1000, 1480, 1.205, 523.7, 572.4, 328.9, 5657.9, 710.5, 750.0, 859.2, 733.6, 1019.7, 425.0, 10, 133.6, 62.5, 442.1, 464.5, 29, 40, 2.175123}
        Tschets(27).Teff = {0.00, 71.0, 80.5, 82.1, 83.3, 84.24, 84.6, 84.04, 81.46, 77.16, 72.0, 60.0}
        Tschets(27).Tverm = {3.0, 9.0, 11.9, 12.7, 13.5, 14.1, 14.8, 15.6, 16.4, 16.9, 17.1, 17.1}
        Tschets(27).TPstat = {3826.1, 4252.2, 4060.9, 3963.5, 3833.9, 3695.7, 3533.9, 3220.0, 2868.7, 2440.9, 2036.5, 1756.5}
        Tschets(27).TPtot = {3826.1, 4289.5, 4149.1, 4072.2, 3967.4, 3854.1, 3719.4, 3457.3, 3164.2, 2804.5, 2463.6, 2308.0}
        Tschets(27).TFlow = {0.00, 1.5, 2.3, 2.55, 2.83, 3.08, 3.34, 3.77, 4.21, 4.67, 5.06, 5.75}
        Tschets(27).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(27).Geljon = {0.14412, 1.3974, -98.38, 0.00040028, 0.19121, -2.6461, 0.009678, 0.032648}

        Tschets(28).Tname = "GALAK"
        Tschets(28).Tdata = {1200, 1465, 1.2, 500, 665, 300, 5500, 820, 886, 1018, 875, 1208, 408, 16, 120.0, 60.0, 455, 500, 38.4, 71.0, 2.100496}
        Tschets(28).Teff = {0.00, 33.1, 52.7, 66.2, 72.2, 74.1, 75.0, 75.1, 74.3, 72.9, 70.5, 63.4}
        Tschets(28).Tverm = {8.5, 15.3, 22.1, 29.1, 35.8, 38.8, 42.0, 45.1, 48.4, 51.5, 54.5, 56.9}
        Tschets(28).TPstat = {5258.2, 5655.1, 5982.7, 6182.0, 6174.5, 6022.2, 5832.8, 5586.9, 5284.4, 4925.3, 4529.3, 4057.0}
        Tschets(28).TPtot = {5258.2, 5670.2, 6043.0, 6317.6, 6415.7, 6327.5, 6209.7, 6043.0, 5827.1, 5562.3, 5268.0, 4905.0}
        Tschets(28).TFlow = {0.00, 1.0, 2.0, 3.0, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5}
        Tschets(28).werkp_opT = {0, 0, 0, 0, 0}

        Tschets(29).Tname = "GW"
        Tschets(29).Tdata = {1000, 1480, 1.205, 127.8, 156.7, 82.5, 4117.5, 618.6, 573.2, 573.2, 536.1, 614.4, 416.5, 12, 20.6, 8.2, 127.8, 167.0, 65, 90, 1.016747}
        Tschets(29).Teff = {0.00, 27.8, 40.2, 43.9, 46.4, 47.5, 47.5, 46.6, 43.8, 39.9, 34.8, 28.3}
        Tschets(29).Tverm = {0.4, 0.7, 1.0, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2.0}
        Tschets(29).TPstat = {3953.0, 4082.0, 4030.0, 3936.0, 3804.0, 3630.0, 3405.0, 3141.0, 2793.0, 2414.0, 1956.0, 1452.0}
        Tschets(29).TPtot = {3953.0, 4089.0, 4058.0, 3978.0, 3863.0, 3710.0, 3509.0, 3273.0, 2953.0, 2595.0, 2189.0, 1727.0}
        Tschets(29).TFlow = {0.00, 0.05, 0.1, 0.13, 0.15, 0.18, 0.2, 0.23, 0.25, 0.28, 0.3, 0.33}
        Tschets(29).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(29).Geljon = {0.1327, 32.1, -31159.0, 0.0001441, 0.26841, -23.368, 0.0004895, 0.001958}

        Tschets(30).Tname = "GWA"
        Tschets(30).Tdata = {1000, 1480, 1.205, 127.8, 156.7, 82.5, 4117.5, 618.6, 1191.8, 573.2, 536.1, 577.3, 492.1, 12, 20.6, 8.2, 127.8, 167.0, 90, 50, 1.016747}
        Tschets(30).Teff = {0.00, 30.27, 38.83, 44.99, 49.14, 51.21, 52.45, 52.1, 50.65, 47.89, 43.59, 38.33}
        Tschets(30).Tverm = {0.4, 0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.2, 1.3, 1.4, 1.5, 1.6}
        Tschets(30).TPstat = {3880.8, 3918.9, 3880.8, 3794.0, 3672.5, 3519.8, 3335.8, 3096.3, 2804.7, 2454.1, 2020.2, 1600.2}
        Tschets(30).TPtot = {3880.8, 3936.3, 3905.1, 3832.2, 3731.5, 3596.1, 3429.5, 3217.8, 2971.3, 2690.1, 2350.0, 1985.5}
        Tschets(30).TFlow = {0.00, 0.05, 0.08, 0.1, 0.13, 0.15, 0.18, 0.2, 0.23, 0.25, 0.28, 0.3}
        Tschets(30).werkp_opT = {0, 0, 0, 0, 0}
        Tschets(30).Geljon = {0.13513, 25.537, -28003.0, 0.0000681, 0.21049, -33.199, 0.000566, 0.001838}

        For j = 1 To 30
            Tschets(j).TFlow_scaled = Tschets(0).TFlow_scaled           '[m3/s]
            Tschets(j).TPstat_scaled = Tschets(0).TPstat_scaled         'Statische druk
            Tschets(j).TPtot_scaled = Tschets(0).TPtot_scaled           'Totale druk
            Tschets(j).Tverm_scaled = Tschets(0).Tverm_scaled           'Rendement[%]
            Tschets(j).Teff_scaled = Tschets(0).Teff_scaled             'Vermogen[kW]
        Next
    End Sub

    Private Sub draw_chart1(Tschets_no As Integer)
        Dim hh As Integer
        Dim debiet As Double
        Dim Q_target, P_target As Double
        Dim Weerstand_Coefficient_line, p_loss_line As Double

        If (Tschets_no < (ComboBox1.Items.Count)) And (Tschets_no >= 0) Then
            Try
                'Clear all series And chart areas so we can re-add them
                Chart1.Series.Clear()
                Chart1.ChartAreas.Clear()
                Chart1.Titles.Clear()

                '----------- Line type-----------
                Chart1.ChartAreas.Add("ChartArea0")
                For hh = 0 To 30
                    Chart1.Series.Add("Series" & hh.ToString)
                    Chart1.Series(hh).ChartArea = "ChartArea0"
                    Chart1.Series(hh).ChartType = SeriesChartType.Line
                    Chart1.Series(hh).SmartLabelStyle.Enabled = True
                Next

                '--------- Legend visible ----------------
                Chart1.Series(4).IsVisibleInLegend = False          'Marker
                For hh = 7 To 30
                    Chart1.Series(hh).IsVisibleInLegend = False     'Vane control lines
                Next

                Chart1.Titles.Add(Tschets(Tschets_no).Tname)
                Chart1.Titles(0).Font = New Font("Arial", 16, System.Drawing.FontStyle.Bold)

                '--------- Legend names ----------------
                Chart1.Series(0).Name = "P totaal [mBar] "
                Chart1.Series(1).Name = "Rendement [%] "
                Chart1.Series(2).Name = "As vermogen [kW] "
                Chart1.Series(3).Name = "Line resistance "
                Chart1.Series(4).Name = "Marker "
                Chart1.Series(5).Name = "P static [mBar] "
                Chart1.Series(6).Name = "Vane-Control "

                Chart1.Series(0).Color = Color.LightGreen   'Total pressure
                Chart1.Series(1).Color = Color.Red          'Efficiency
                Chart1.Series(2).Color = Color.LightGreen   'Power
                Chart1.Series(3).Color = Color.Blue         'marker
                Chart1.Series(4).Color = Color.LightBlue    'Line resistance
                Chart1.Series(5).Color = Color.Blue         'Static pressure

                '----------- labels on-off ------------------       
                If CheckBox6.Checked Then   'Labels on
                    Chart1.Series(0).IsValueShownAsLabel = True
                    Chart1.Series(1).IsValueShownAsLabel = True
                    Chart1.Series(2).IsValueShownAsLabel = True
                    Chart1.Series(5).IsValueShownAsLabel = True
                End If

                '------- line thickness-------------
                For hh = 0 To 30
                    Chart1.Series(hh).BorderWidth = 1
                Next
                Chart1.Series(5).BorderWidth = 3    'Static pressure

                Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
                Chart1.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
                Chart1.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True
                Chart1.ChartAreas("ChartArea0").AxisY2.MinorTickMark.Enabled = True

                '------------- Grid --------------------------
                'Chart1.ChartAreas("ChartArea0").AxisY.MajorGrid.Enabled = True
                'Chart1.ChartAreas("ChartArea0").AxisX.MajorGrid.Enabled = True
                'Chart1.ChartAreas("ChartArea0").AxisY.MinorGrid.Enabled = True
                'Chart1.ChartAreas("ChartArea0").AxisX.MinorGrid.Enabled = True

                '---------------- fan target info---------------------
                TextBox149.Text = Round(G_Debiet_z_act_hr, 0).ToString  'Debiet [Am3/hr]
                TextBox148.Text = NumericUpDown37.Value.ToString        'Pstatic [mbar]
                TextBox156.Text = NumericUpDown12.Value.ToString        'Density [kg/m3]

                Chart1.ChartAreas("ChartArea0").AxisY.Title = "Ptotaal [mBar]"
                Chart1.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical
                Chart1.ChartAreas("ChartArea0").AxisY2.Enabled = AxisEnabled.True
                Chart1.ChartAreas("ChartArea0").AxisY2.Title = "Rendement [%] As-vermogen [kW]"


                '-------- Pressure on left hand Y axis------
                For hh = 0 To 20
                    Chart1.Series(hh).YAxisType = AxisType.Primary
                Next

                '-------- Power on right hand Y axis------
                Chart1.Series(1).YAxisType = AxisType.Secondary
                Chart1.Series(2).YAxisType = AxisType.Secondary
                For hh = 20 To 30
                    Chart1.Series(hh).YAxisType = AxisType.Secondary
                Next

                '------------------ Grafiek tekst en target ---------------------
                If CheckBox2.Checked Then               '========Per uur=========
                    Q_target = G_Debiet_z_act_hr                                            '[Am3/hr]
                    P_target = NumericUpDown37.Value                                        '[mBar] Gewenste fan  gegevens
                    Chart1.ChartAreas("ChartArea0").AxisX.Title = "Debiet [Am3/hr]"
                Else                                    '========Per seconde=========
                    Q_target = G_Debiet_z_act_hr / 3600                                     '[Am3/sec]
                    P_target = NumericUpDown37.Value                                        '[mBar] Gewenste fan  gegevens
                    Chart1.ChartAreas("ChartArea0").AxisX.Title = "Debiet [Am3/sec]"
                End If

                '---------------- add the fan lines-----------------------
                For hh = 0 To 50
                    debiet = case_x_flow(hh, 0)
                    If CheckBox2.Checked Then debiet = Round(debiet * 3600, 1)      'Per uur
                    Chart1.Series(0).Points.AddXY(debiet, Round(case_x_Ptot(hh), 1))
                    Chart1.Series(5).Points.AddXY(debiet, Round(case_x_Pstat(hh, 0), 1))
                    If CheckBox7.Checked Then Chart1.Series(1).Points.AddXY(debiet, Round(case_x_Efficiency(hh), 1))    'efficiency
                    If CheckBox8.Checked Then Chart1.Series(2).Points.AddXY(debiet, Round(case_x_Power(hh, 0), 1))      'Power
                Next hh

                '----------- adding labels ----------------
                Chart1.Series(0).Points(45).Label = "P.total"
                Chart1.Series(5).Points(45).Label = "P.static"
                If CheckBox7.Checked Then Chart1.Series(1).Points(45).Label = "Efficiency"  'efficiency
                If CheckBox8.Checked Then Chart1.Series(2).Points(45).Label = "Power"       'Power

                '-------------------Target dot ---------------------
                If CheckBox3.Checked Then
                    Chart1.Series(3).YAxisType = AxisType.Primary
                    Chart1.Series(3).Points.AddXY(Q_target, P_target)
                    Chart1.Series(3).Points(0).MarkerStyle = DataVisualization.Charting.MarkerStyle.Star10
                    Chart1.Series(3).Points(0).MarkerSize = 20

                    '---------------- add the Duct resistance line-----------------------
                    Weerstand_Coefficient_line = P_target * 2 / (NumericUpDown12.Value * Q_target ^ 2)
                    For hh = 0 To 50
                        debiet = case_x_flow(hh, 0)
                        If CheckBox2.Checked Then debiet = Round(debiet * 3600, 1)      'Per uur

                        If CheckBox10.Checked Then
                            p_loss_line = 0.5 * Weerstand_Coefficient_line * NumericUpDown12.Value * debiet ^ 2
                            If p_loss_line < P_target * 1.1 Then Chart1.Series(4).Points.AddXY(debiet, p_loss_line)
                        End If
                    Next
                End If

                '-------------------Inlet Vane Control lines ---------------------
                If CheckBox13.Checked Then
                    Dim VC_phi = New Double() {0.02, 0.04, 0.06, 0.08, 0.1, 0.12}     'Pressure loss coeff (*= 2.5)
                    Dim VC_open = New String() {"80", "70", "60", "50", "40", "30"}
                    Dim point_count As Integer
                    For jj = 0 To 5
                        For hh = 0 To 50
                            debiet = case_x_flow(hh, 0)
                            If CheckBox2.Checked Then debiet = Round(debiet * 3600, 1)      'Per uur
                            P_loss_IVC(jj + 10, VC_phi(jj), hh, debiet, VC_open(jj))         'Calc and plot to chart
                        Next
                        point_count = Chart1.Series(jj + 10).Points.Count - 1                'Last plotted point
                        Chart1.Series(jj + 10).Points(point_count).Label = VC_open(jj) & "°" 'Add the VC opening angle 
                    Next
                End If

                '---------------Outlet damper-----------------
                If CheckBox14.Checked Then
                    ' P_loss_Out_Flow_damper(Series As Integer, phi As Double, hh As Integer, debiet As Double, alpha As Double)
                End If

                Chart1.Refresh()
            Catch ex As Exception
                'MessageBox.Show(ex.Message & "Line 1780")  ' Show the exception's message.
            End Try
        End If
    End Sub
    'Pressure loss over the Inlet Vane control
    Private Sub P_loss_IVC(series As Integer, phi As Double, hh As Integer, debiet As Double, alpha As Double)
        Dim ivc_area, ivc_speed, ivc_dia, ivc_Power, ivc_loss As Double
        Dim fan_P_static, Pstatic_w_ivc, debiet_sec As Double

        Chart1.Series(series).Color = Color.Black                           'Static pressure

        Double.TryParse(TextBox159.Text, ivc_dia)
        ivc_dia /= 1000                                                     '[m] diameter
        ivc_area = 3.14 / 4 * ivc_dia ^ 2 * Sin(alpha * PI / 180)           '[m2] open area

        debiet_sec = debiet                                                 'debiet in [m3/sec]
        If CheckBox2.Checked Then debiet_sec = debiet / 3600                'debiet in [m3/hr]
        ivc_speed = debiet_sec / ivc_area                                   '[m/s]
        ivc_loss = 0.5 * phi * NumericUpDown12.Value * ivc_speed ^ 2        '[mbar]
        fan_P_static = case_x_Pstat(hh, 0)
        Pstatic_w_ivc = fan_P_static - ivc_loss                             '[mbar] Pstatic with IVC
        If Pstatic_w_ivc < 0 Then Pstatic_w_ivc = 0
        If ivc_loss < fan_P_static * 0.8 Then Chart1.Series(series).Points.AddXY(debiet, Round(Pstatic_w_ivc, 1))

        '----------- Power lines IVC ------------
        If CheckBox14.Checked Then
            ivc_Power = 15 * Pstatic_w_ivc * debiet / case_x_Efficiency(hh)  '[kW]
            If hh > 10 And ivc_Power > 1 Then Chart1.Series(series + 10).Points.AddXY(debiet, Round(ivc_Power, 1))
        End If
    End Sub
    'Pressure loss over the Outlet Flow Damper
    Private Sub P_loss_Out_Flow_damper(series As Integer, phi As Double, hh As Integer, debiet As Double, alpha As Double)
        'Future
        'Future
        'Future
    End Sub

    'Write to Word document
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        write_to_word()
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Selectie_1()
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Selectie_1()
    End Sub
    'Calculatie soortelijk gewicht
    'Zonder vochtigheidcompensatie
    Private Function calc_sg_air(P As Double, T As Double, RH As Double, MG As Double)
        Dim p1, p2, sg1, sg2 As Double

        Select Case True
            Case RadioButton1.Checked                       'Medium is Lucht
                NumericUpDown8.Value = 28.96                'According ISO6972 for dry air
                Label96.Text = "Air"
                NumericUpDown8.BackColor = Color.White
            Case RadioButton2.Checked                       'Medium is gas, mol weight is entered
                Label96.Text = "Enter mol whgt"
                NumericUpDown8.BackColor = Color.Yellow
        End Select

        '---------------- We assume that above the 100c the air is dry ----------------------------
        '--------------- otherwise unpredictable results-------------------------------------------
        If T > 100 Then
            NumericUpDown5.Value = 0
        End If

        '--------------------------------Partiele waterdamp druk----------------------------
        If T >= 0 And T <= 99 Then
            p1 = Pow(10, (8.07131 - (1730.63 / (233.462 + T)))) * 133.322368
        End If

        If T >= 100 And T < 374 Then
            p1 = Pow(10, (8.14019 - (1810.94 / (244.485 + T)))) * 133.322368
        End If

        If RadioButton1.Checked Or RadioButton2.Checked Then
            RH = 0              'RH is not calculated
            p1 = 0              'Partial pressure water not calculated    
        End If

        p1 = p1 * RH / 100
        p2 = P - p1

        '-------------------------------- soortelijk gewicht---------------------------
        'gecontroleerd tegen http://www.denysschen.com/catalogue/density.aspx --------
        'Algmene Gasconstante is 8314,32
        '8314,32/28.8= 288.69


        sg1 = p1 / (461.495 * (T + 273.15))             'Water vapor
        sg2 = p2 / (8.31432 / MG * (T + 273.15))        'Droge lucht

        Return (sg1 + sg2)
    End Function
    'Save data and line chart to file
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Update_selectie()
        Calc_noise()
        write_to_word()
    End Sub
    'Find workpoint hight efficiency
    Private Sub Find_hi_eff()
        Dim hh, jj, pos_counter As Integer
        Dim eff_hi As Double

        For jj = 0 To (Tschets.Length - 2)                '30 T schetsen 
            eff_hi = 0
            pos_counter = 0
            For hh = 0 To 11                'Check all Efficiencies to find the highest
                If Tschets(jj).Teff(hh) > eff_hi Then
                    eff_hi = Tschets(jj).Teff(hh)
                    pos_counter = hh
                End If
            Next hh

            Tschets(jj).werkp_opT(0) = Tschets(jj).Teff(pos_counter)    'rendement
            Tschets(jj).werkp_opT(1) = Tschets(jj).TPtot(pos_counter)   'P_totaal [Pa]
            Tschets(jj).werkp_opT(2) = Tschets(jj).TPstat(pos_counter)  'P_statisch [Pa]
            Tschets(jj).werkp_opT(3) = Tschets(jj).Tverm(pos_counter)   'as_vermogen [kW]
            Tschets(jj).werkp_opT(4) = Tschets(jj).TFlow(pos_counter)   'debiet[m3/sec]
            'MessageBox.Show("JJ=" & jj.ToString &" aantal hh =" & hh)
        Next jj
    End Sub

    'Write data to Word 
    Private Sub write_to_word()
        Dim bmp_tab_page1 As New Bitmap(TabPage1.Width, TabPage1.Height)
        Dim bmp_grouobox23 As New Bitmap(GroupBox23.Width, GroupBox23.Height)
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara4 As Word.Paragraph

        'Start Word and open the document template. 
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering department"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = 16
        oPara1.Range.Font.Bold = True
        oPara1.Format.SpaceAfter = 4                '4 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = 11
        oPara2.Format.SpaceAfter = 2
        oPara2.Range.Font.Bold = False
        oPara2.Range.Text = "Fan selection and sizing " & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 11
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        oTable.Cell(1, 1).Range.Text = "Project Name"
        oTable.Cell(1, 2).Range.Text = TextBox283.Text
        oTable.Cell(2, 1).Range.Text = "Item number"
        oTable.Cell(2, 2).Range.Text = TextBox284.Text
        oTable.Cell(3, 1).Range.Text = "Fan type "
        oTable.Cell(3, 2).Range.Text = Label1.Text
        oTable.Cell(4, 1).Range.Text = "Author "
        oTable.Cell(4, 2).Range.Text = Environment.UserName
        oTable.Cell(5, 1).Range.Text = "Date "
        oTable.Cell(6, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(2)

        oTable.Rows.Item(1).Range.Font.Bold = True
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 20 x 10 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 23, 10)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 9
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        For j = 0 To 23 'Rows
            oTable.Cell(j + 1, 1).Range.Text = case_x_conditions(j, 10)     'Write all variables
            oTable.Cell(j + 1, 2).Range.Text = case_x_conditions(j, 11)     'Write all units
            oTable.Cell(j + 1, 3).Range.Text = case_x_conditions(j, 1)      'Case 1
            oTable.Cell(j + 1, 4).Range.Text = case_x_conditions(j, 2)      'Case 2
            oTable.Cell(j + 1, 5).Range.Text = case_x_conditions(j, 3)      'Case 3
            oTable.Cell(j + 1, 6).Range.Text = case_x_conditions(j, 4)      'Case 4
            oTable.Cell(j + 1, 7).Range.Text = case_x_conditions(j, 5)      'Case 5
            oTable.Cell(j + 1, 8).Range.Text = case_x_conditions(j, 6)      'Case 6
            oTable.Cell(j + 1, 9).Range.Text = case_x_conditions(j, 7)      'Case 5
            oTable.Cell(j + 1, 10).Range.Text = case_x_conditions(j, 8)     'Case 6
        Next

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(1.3)   'Change width of columns 
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.75)
        oTable.Columns.Item(4).Width = oWord.InchesToPoints(0.75)
        oTable.Columns.Item(5).Width = oWord.InchesToPoints(0.75)
        oTable.Columns.Item(6).Width = oWord.InchesToPoints(0.45)
        oTable.Columns.Item(7).Width = oWord.InchesToPoints(0.45)
        oTable.Columns.Item(8).Width = oWord.InchesToPoints(0.45)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '------------------ Noise Details----------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 9
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True
        oTable.Cell(1, 1).Range.Text = "Noise"
        oTable.Cell(2, 1).Range.Text = "Sound power Discharge (induct) "
        oTable.Cell(2, 2).Range.Text = TextBox361.Text
        oTable.Cell(2, 3).Range.Text = "[dBA], with " & ComboBox5.SelectedItem
        oTable.Cell(3, 1).Range.Text = "Sound power Suction (induct) "
        oTable.Cell(3, 2).Range.Text = TextBox333.Text
        oTable.Cell(3, 3).Range.Text = "[dBA] with, " & ComboBox10.SelectedItem
        oTable.Cell(4, 1).Range.Text = "Sound pressure Casing"
        oTable.Cell(4, 2).Range.Text = TextBox296.Text
        oTable.Cell(4, 3).Range.Text = "[dBA] @ 1m, with " & ComboBox9.SelectedItem
        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(3.2)
        oTable.Rows.Item(1).Range.Font.Bold = True
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '------------------ motor----------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 9
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True
        oTable.Cell(1, 1).Range.Text = "Electric Motor"
        oTable.Cell(2, 1).Range.Text = "Speed"
        oTable.Cell(2, 2).Range.Text = " "
        oTable.Cell(2, 3).Range.Text = "[rpm]"
        oTable.Cell(3, 1).Range.Text = "Power"
        oTable.Cell(3, 2).Range.Text = ""
        oTable.Cell(3, 3).Range.Text = "[kW]"
        oTable.Columns.Item(1).Width = oWord.InchesToPoints(1.3)   'Change width of columns
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(1.55)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.8)
        oTable.Rows.Item(1).Range.Font.Bold = True
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


        '------------------save Chart1 ---------------- 
        Scale_rules_applied(ComboBox1.SelectedIndex, NumericUpDown9.Value, NumericUpDown10.Value, NumericUpDown12.Value)
        draw_chart1(ComboBox1.SelectedIndex)
        Chart1.SaveImage("c:\Temp\Chart1.Jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)
        oPara4 = oDoc.Content.Paragraphs.Add
        oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oPara4.Range.InlineShapes.AddPicture("c:\Temp\Chart1.Jpeg")
        oPara4.Range.InlineShapes.Item(1).LockAspectRatio = True
        oPara4.Range.InlineShapes.Item(1).Width = 500
        oPara4.Range.InsertParagraphAfter()

        '------------------save Chart5 ---------------- 
        draw_chart5()
        Chart5.SaveImage("c:\Temp\Chart5.Jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)
        oPara4 = oDoc.Content.Paragraphs.Add
        oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oPara4.Range.InlineShapes.AddPicture("c:\Temp\Chart5.Jpeg")
        oPara4.Range.InlineShapes.Item(1).LockAspectRatio = True
        oPara4.Range.InlineShapes.Item(1).Width = 500
        oPara4.Range.InsertParagraphAfter()


        '---- save tab page 1---------------
        TabPage2.Show()
        TabPage2.Refresh()
        TabPage1.DrawToBitmap(bmp_tab_page1, DisplayRectangle)
        bmp_tab_page1.Save("c:\Temp\page2.Jpeg", Imaging.ImageFormat.Jpeg)
        oPara4 = oDoc.Content.Paragraphs.Add
        oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oPara4.Range.InlineShapes.AddPicture("c:\Temp\page2.Jpeg")
        oPara4.Range.InlineShapes.Item(1).LockAspectRatio = True
        'oPara4.Range.InlineShapes.Item(1).Width = 400
        oPara4.Range.InsertParagraphAfter()

        '---- save geluid page 1---------------
        TabPage4.Show()
        TabPage4.Refresh()
        GroupBox23.DrawToBitmap(bmp_grouobox23, DisplayRectangle)
        bmp_grouobox23.Save("c:\Temp\page3.Jpeg", Imaging.ImageFormat.Jpeg)
        oPara4 = oDoc.Content.Paragraphs.Add
        oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oPara4.Range.InlineShapes.AddPicture("c:\Temp\page3.Jpeg")
        oPara4.Range.InlineShapes.Item(1).LockAspectRatio = True
        'oPara4.Range.InlineShapes.Item(1).Width = 400
        oPara4.Range.InsertParagraphAfter()

    End Sub
    'Graphic next model
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox1.SelectedIndex < (ComboBox1.Items.Count - 1) Then
            ComboBox1.SelectedIndex += 1
            Scale_rules_applied(ComboBox1.SelectedIndex, NumericUpDown9.Value, NumericUpDown10.Value, NumericUpDown12.Value)
            draw_chart1(ComboBox1.SelectedIndex)
        End If
    End Sub
    'Graphic previous model
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If (ComboBox1.SelectedIndex > 0) Then
            ComboBox1.SelectedIndex -= 1
            Scale_rules_applied(ComboBox1.SelectedIndex, NumericUpDown9.Value, NumericUpDown10.Value, NumericUpDown12.Value)
            draw_chart1(ComboBox1.SelectedIndex)
        End If
    End Sub
    'Scale rules capacity
    '1= inlet, 2=outlet
    'Diameter in [m]
    'Q=Capacity in [m3/s]
    'Speed in [rpm] or [rad/s] or [rps] if used consequently
    'Note; sg medum speelt geen rol !!!!!!!!!!
    Private Function Scale_rule_cap(QQ1 As Double, Dia1 As Double, Dia2 As Double, n1 As Double, n2 As Double)
        Dim QQ2 As Double

        QQ2 = QQ1 * (n2 / n1) * (Dia2 / Dia1) ^ 3

        Return (QQ2)
    End Function

    'Scale rules Pressure, Total and Static
    'Pressure in [Pa] or [Pa abs] or [wwmc] if used consequently
    '1= inlet, 2=outlet
    'Diameter in [m]
    'Q=Capacity in [m3/s]
    'Speed in [rpm] or [rad/s] or [rps] if used consequently
    Private Function Scale_rule_Pressure(Pt1 As Double, Dia1 As Double, Dia2 As Double, n1 As Double, n2 As Double, Ro1 As Double, Ro2 As Double)
        Dim Pt2 As Double

        Pt2 = Pt1 * (n2 / n1) ^ 2 * (Ro2 / Ro1) * (Dia2 / Dia1) ^ 2

        Return (Pt2)
    End Function
    'Scale rules Power
    '1= old, 2= new
    'Diameter in [m]
    'Q=Capacity in [m3/s]
    'Speed in [rpm] or [rad/s] or [rps] if used consequently
    Private Function Scale_rule_Power(Power1 As Double, Dia1 As Double, Dia2 As Double, n1 As Double, n2 As Double, Ro1 As Double, Ro2 As Double)
        Dim Power2 As Double

        Power2 = Power1 * (n2 / n1) ^ 3 * (Ro2 / Ro1) * (Dia2 / Dia1) ^ 5

        Return (Power2)
    End Function

    Private Sub Chart1_Layout(sender As Object, e As LayoutEventArgs) Handles Chart1.Layout
        Dim ww, hh, sstep As Integer

        ww = Chart1.Size.Width - 170
        hh = Chart1.Size.Height * 0.1 + 100
        sstep = 6

        GroupBox36.Location = New Point(ww, hh)
        hh += GroupBox36.Height + sstep
        GroupBox37.Location = New Point(ww, hh)
        hh += GroupBox37.Height + sstep
        GroupBox39.Location = New Point(ww, hh)
        hh += GroupBox39.Height + sstep
        GroupBox38.Location = New Point(ww, hh)
        hh += GroupBox38.Height + sstep
        GroupBox42.Location = New Point(ww, hh)
        hh += GroupBox42.Height + sstep
        Button3.Location = New Point(ww, hh)  'Save button

        Label66.Location = New Point(10, 10)
    End Sub
    'Diameter or speed impeller changed or density changed, recalculate and draw chart
    Private Sub Scale_rules_applied(ty As Integer, Dia1 As Double, n2 As Double, Ro1 As Double)
        Dim hh As Integer

        If (ty >= 0) And (ty < (ComboBox1.Items.Count)) Then    'Preventing exceptions
            Try
                For hh = 0 To 11                                '10 Meetpunten per Tschets type
                    If Tschets(ty).TPtot(hh) > 0 Then           'Rest of the array is empty
                        '============================ TSchets data inlezen ================================
                        cond(4).Typ = ty                                                '[-]        T_SCHETS  
                        cond(4).Q0 = Tschets(ty).TFlow(hh)                              '[Am3/s]    T_SCHETS Volume debiet
                        cond(4).Pt0 = Tschets(ty).TPtot(hh)                             '[PaG]      T_SCHETS Ptotal          
                        cond(4).Ps0 = Tschets(ty).TPstat(hh)                            '[PaG]      T_SCHETS Pstatic 
                        cond(4).Power0 = Tschets(ty).Tverm(hh)                          '[PaG]      T_SCHETS vermogen 
                        cond(4).Rpm0 = Tschets(ty).Tdata(1)                             '[rpm]      T_SCHETS
                        cond(4).Dia0 = Tschets(ty).Tdata(0)                             '[mm]       T_SCHETS diameter waaier
                        cond(4).Ro0 = Tschets(cond(1).Typ).Tdata(2)                     '[kg/m3]    T_SCHETS density inlet flange                         

                        '======================== data van het scherm inlezen ============================
                        cond(4).Dia1 = Dia1                                             '[mm]           
                        cond(4).Qkg = NumericUpDown3.Value                              '[kg/hr]
                        cond(4).Pt1 = NumericUpDown76.Value * 100                       'Press total [Pa] abs. inlet flange                                           
                        cond(4).Ps1 = NumericUpDown76.Value * 100                       'Press total [Pa] abs. inlet flange                                 
                        cond(4).Rpm1 = n2                                               '[rpm]          
                        cond(4).Qkg = NumericUpDown3.Value                              '[kg/hr]       
                        cond(4).Ro1 = NumericUpDown12.Value                             '[kg/m3] density inlet flange       
                        cond(4).T1 = NumericUpDown4.Value                               '[c]                                

                        '======================== waaier #1 ===============================================
                        Calc_stage(cond(4))                     'Bereken de waaier #1  (in elf stappen)
                        calc_loop_loss(cond(4))                 'Bereken de omloop verliezen 

                        '======================== waaier #2 ===============================================
                        cond(5) = cond(4)                       'Kopieer de struct met gegevens
                        cond(5).T1 = cond(4).T2                 '[c] uitlaat waaier#1 is inlaat waaier #2
                        cond(5).Pt1 = cond(4).Ps3               'Inlaat waaier #2 
                        cond(5).Ps1 = cond(4).Ps3               'Inlaat waaier #2
                        cond(5).Ro1 = cond(4).Ro3               'Ro Inlaat waaier #2

                        Calc_stage(cond(5))                     'Bereken de waaier #2  
                        calc_loop_loss(cond(5))                 'Bereken de omloop verliezen  

                        '======================== waaier #3 ===============================================
                        cond(6) = cond(5)                       'Kopieer de struct met gegevens
                        cond(6).T1 = cond(5).T2                 '[c] uitlaat waaier#1 is inlaat waaier #2
                        cond(6).Pt1 = cond(5).Ps3               'Inlaat waaier #2 
                        cond(6).Ps1 = cond(5).Ps3               'Inlaat waaier #2
                        cond(6).Ro1 = cond(5).Ro3               'Ro Inlaat waaier #2

                        Calc_stage(cond(6))                     'Bereken de waaier #2  
                        calc_loop_loss(cond(6))                 'Bereken de omloop verliezen (niet echt nodig)

                        Select Case True
                            Case RadioButton9.Checked      '1 traps
                                Tschets(ty).TFlow_scaled(hh) = cond(4).Q1                                       '[Am3/hr]
                                Tschets(ty).TPtot_scaled(hh) = Round((cond(4).Pt2 - cond(4).Pt1) / 100, 4)      '[mbar] dP fan total
                                Tschets(ty).TPstat_scaled(hh) = Round((cond(4).Ps2 - cond(4).Ps1) / 100, 4)     '[mbar] dP fan static
                                Tschets(ty).Tverm_scaled(hh) = Round(cond(4).Power, 4)                          '[kW]
                                Tschets(ty).Teff_scaled(hh) = Round((100 * cond(4).delta_pt * cond(4).Q1 / (Tschets(ty).Tverm_scaled(hh) * 1000)), 4)

                            Case RadioButton10.Checked      '2 traps
                                Tschets(ty).TFlow_scaled(hh) = cond(4).Q1                                        '[Am3/hr]
                                Tschets(ty).TPtot_scaled(hh) = Round((cond(5).Pt2 - cond(4).Pt1) / 100, 4)       '[mbar] dP fan total
                                Tschets(ty).TPstat_scaled(hh) = Round((cond(5).Ps2 - cond(4).Ps1) / 100, 4)      '[mbar] dP fan static
                                Tschets(ty).Tverm_scaled(hh) = Round(cond(4).Power + cond(5).Power, 4)           '[kW] waaier 1+2
                                Tschets(ty).Teff_scaled(hh) = Round((100 * cond(5).delta_pt * cond(5).Q1 / (Tschets(ty).Tverm_scaled(hh) * 1000)), 4)

                            Case RadioButton11.Checked   '3 traps
                                Tschets(ty).TFlow_scaled(hh) = cond(4).Q1                                               '[Am3/hr]
                                Tschets(ty).TPtot_scaled(hh) = Round((cond(6).Pt2 - cond(4).Pt1) / 100, 4)              '[mbar] dP fan total
                                Tschets(ty).TPstat_scaled(hh) = Round((cond(6).Ps2 - cond(4).Ps1) / 100, 4)             '[mbar] dP fan static
                                Tschets(ty).Tverm_scaled(hh) = Round(cond(4).Power + cond(5).Power + cond(6).Power, 4)  '[kW] waaier 1+2+3
                                Tschets(ty).Teff_scaled(hh) = Round((100 * cond(6).delta_pt * cond(6).Q1 / (Tschets(ty).Tverm_scaled(hh) * 1000)), 4)

                        End Select
                    End If
                Next hh


                '-----------------------------------------------------------
                ' From the Tschets data points the polynoom variable are calculated
                ' Then 50 chart points are calculated to get a smooth line
                '-----------------------------------------------------------
                Dim j As Integer
                Dim t() As PPOINT
                Dim flow, max_flow As Double
                Dim aantal_Tschets_punten As Integer = 10
                Dim aantal_Grafiek_punten As Integer = 50

                ReDim PZ(aantal_Tschets_punten)          'PZ() too big gives wrong results !!!
                max_flow = Tschets(ty).TFlow_scaled(aantal_Tschets_punten)

                '=============== calculate the polynoom for Ptotal ====================
                For j = 0 To aantal_Tschets_punten       'Get data into PZ()
                    PZ(j).x = Tschets(ty).TFlow_scaled(j)
                    PZ(j).y = Tschets(ty).TPtot_scaled(j)
                Next
                t = Trend(PZ, 5)            'calculate the polynoom

                For j = 0 To 5      'Store the variable for later use
                    ABCDE_Ptot(j) = BZ(j, 0)
                Next


                For j = 0 To aantal_Grafiek_punten      'Calculate chart data points
                    flow = j / aantal_Grafiek_punten * max_flow
                    case_x_flow(j, 0) = flow
                    case_x_Ptot(j) = BZ(0, 0) + BZ(1, 0) * flow ^ 1 + BZ(2, 0) * flow ^ 2 + BZ(3, 0) * flow ^ 3 + BZ(4, 0) * flow ^ 4 + BZ(5, 0) * flow ^ 5
                Next

                '=============== calculate the polynoom for Pstatic ====================
                For j = 0 To aantal_Tschets_punten       'Get data into PZ()
                    PZ(j).x = Tschets(ty).TFlow_scaled(j)
                    PZ(j).y = Tschets(ty).TPstat_scaled(j)
                Next
                t = Trend(PZ, 5)        'calculate the polynoom

                For j = 0 To 5      'Store the variable for later use
                    ABCDE_Psta(j) = BZ(j, 0)
                Next


                For j = 0 To aantal_Grafiek_punten      'Calculate chart data points
                    flow = j / aantal_Grafiek_punten * max_flow
                    case_x_flow(j, 0) = flow
                    case_x_Pstat(j, 0) = BZ(0, 0) + BZ(1, 0) * flow ^ 1 + BZ(2, 0) * flow ^ 2 + BZ(3, 0) * flow ^ 3 + BZ(4, 0) * flow ^ 4 + BZ(5, 0) * flow ^ 5
                Next

                ''=============== calculate the polynoom for Power ====================
                For j = 0 To aantal_Tschets_punten       'Get data into PZ()
                    PZ(j).x = Tschets(ty).TFlow_scaled(j)
                    PZ(j).y = Tschets(ty).Tverm_scaled(j)
                Next
                t = Trend(PZ, 5)        'calculate the polynoom

                For j = 0 To 5      'Store the variable for later use
                    ABCDE_Pow(j) = BZ(j, 0)
                Next

                For j = 0 To aantal_Grafiek_punten      'Calculate chart data points
                    flow = j / aantal_Grafiek_punten * max_flow
                    case_x_flow(j, 0) = flow
                    case_x_Power(j, 0) = BZ(0, 0) + BZ(1, 0) * flow ^ 1 + BZ(2, 0) * flow ^ 2 + BZ(3, 0) * flow ^ 3 + BZ(4, 0) * flow ^ 4 + BZ(5, 0) * flow ^ 5
                Next

                '=============== convert to polynoom, Efficiency ====================
                TextBox158.AppendText(Environment.NewLine)
                For j = 0 To aantal_Tschets_punten       'Get data into PZ()
                    PZ(j).x = Tschets(ty).TFlow_scaled(j)
                    PZ(j).y = Tschets(ty).Teff_scaled(j)
                Next
                t = Trend(PZ, 5)        'calculate the polynoom

                For j = 0 To 5      'Store the variable for later use
                    ABCDE_Eff(j) = BZ(j, 0)
                Next

                For j = 0 To aantal_Grafiek_punten      'Calculate chart data points
                    flow = j / aantal_Grafiek_punten * max_flow
                    case_x_flow(j, 0) = flow
                    case_x_Efficiency(j) = BZ(0, 0) + BZ(1, 0) * flow ^ 1 + BZ(2, 0) * flow ^ 2 + BZ(3, 0) * flow ^ 3 + BZ(4, 0) * flow ^ 4 + BZ(5, 0) * flow ^ 5
                Next

                draw_chart1(ty)
            Catch ex As Exception
                MessageBox.Show(ex.Message & " Problem in Scale_rules_applied, line 2033")  ' Show the exception's message.
            End Try
        End If
    End Sub

    Private Sub do_Chart2()
        Dim schets_no As Integer

        Chart2.Series.Clear()
        Chart2.ChartAreas.Clear()
        Chart2.Titles.Clear()
        Chart2.ChartAreas.Add("ChartArea0")

        For schets_no = 0 To (ComboBox1.Items.Count - 1)
            Chart2.Series.Add("Series" & schets_no.ToString)
            Chart2.Series(schets_no).ChartArea = "ChartArea0"
            Chart2.Series(schets_no).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart2.Series(schets_no).Name = (Tschets(schets_no).Tname)
            Chart2.Series(schets_no).BorderWidth = 1
        Next

        Chart2.Series(ComboBox2.SelectedIndex).BorderWidth = 4
        Chart2.Series(ComboBox2.SelectedIndex).Color = Color.Red

        Chart2.Titles.Add("Performance volgens de T-schetsen @ dia= 1000mm, ro= 1.20 kg/m3, 1480 rpm")
        Chart2.ChartAreas("ChartArea0").AxisX.Title = "Debiet [Am3/s]"
        Chart2.ChartAreas("ChartArea0").AxisY.Title = "Ptotaal [Pa]"

        For schets_no = 0 To (ComboBox1.Items.Count - 1)    'Fill line chart
            Scale_rules_applied(schets_no, 1000, 1480, 1.2) 'Compare alle fans on the same basis
            For hh = 0 To 11   'Fill line chart
                If Tschets(schets_no).TPtot(hh) > 0 Then           'Rest of the array is empty
                    Chart2.Series(schets_no).Points.AddXY(Tschets(schets_no).TFlow_scaled(hh), Tschets(schets_no).TPtot_scaled(hh))
                End If
            Next hh
        Next schets_no
    End Sub
    Private Sub TabPage7_Paint(sender As Object, e As PaintEventArgs) Handles TabPage7.Paint
        do_Chart2()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged, NumericUpDown18.ValueChanged
        Dim qq, sigma02 As Double

        If (ComboBox3.SelectedIndex > -1) Then      'Prevent exceptions
            Dim words() As String = steel(ComboBox3.SelectedIndex).Split(";")
            TextBox33.Text = words(6)     'Density steel
            Label106.Text = words(20)     'Opmerkingen

            '--------------- select the strength @ temperature
            qq = NumericUpDown18.Value

            Select Case True
                Case (qq > -10 AndAlso qq <= 0)
                    Double.TryParse(words(9), sigma02)     'Sigma 0.2 [N/mm]
                Case (qq > 0 AndAlso qq <= 20)
                    Double.TryParse(words(10), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 20 AndAlso qq <= 50)
                    Double.TryParse(words(11), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 50 AndAlso qq <= 100)
                    Double.TryParse(words(12), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 100 AndAlso qq <= 150)
                    Double.TryParse(words(13), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 150 AndAlso qq <= 200)
                    Double.TryParse(words(13), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 200 AndAlso qq <= 250)
                    Double.TryParse(words(14), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 250 AndAlso qq <= 300)
                    Double.TryParse(words(15), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 300 AndAlso qq <= 350)
                    Double.TryParse(words(16), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 350 AndAlso qq <= 400)
                    Double.TryParse(words(17), sigma02)    'Sigma 0.2 [N/mm]
            End Select
            TextBox34.Text = sigma02.ToString
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown27.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown22.ValueChanged, TabPage5.Enter, NumericUpDown16.ValueChanged, ComboBox4.SelectedIndexChanged, RadioButton7.CheckedChanged, NumericUpDown15.ValueChanged, NumericUpDown46.ValueChanged, NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown48.ValueChanged, NumericUpDown56.ValueChanged, NumericUpDown49.ValueChanged, RadioButton15.CheckedChanged
        calc_bearing_belts()
        Torsional_stiffness()
        Torsional_analyses()
    End Sub
    'Stiffness calculation of the impellear shaft
    Private Sub Torsional_stiffness()
        Dim g_modulus As Double             'modulus Of elasticity In shear
        Dim k_total, k2 As Double           'modulus Of elasticity In shear

        TextBox257.Text = Round(NumericUpDown44.Value * PI / 180, 0).ToString       'Coupling, Convert rad to degree

        g_modulus = 79.3 * 10 ^ 9           '[Pa] (kilo, mega, giga) steel shear modulus !

        section(0).dia = NumericUpDown25.Value / 1000
        section(0).length = NumericUpDown22.Value / 1000
        section(0).k_stiffness = PI * g_modulus * section(0).dia ^ 4 / (32 * section(0).length)     '[Nm/rad]

        section(1).dia = NumericUpDown26.Value / 1000
        section(1).length = NumericUpDown23.Value / 1000
        section(1).k_stiffness = PI * g_modulus * section(1).dia ^ 4 / (32 * section(1).length)

        section(2).dia = NumericUpDown27.Value / 1000
        section(2).length = NumericUpDown24.Value / 1000
        section(2).k_stiffness = PI * g_modulus * section(2).dia ^ 4 / (32 * section(2).length)

        section(3).dia = NumericUpDown49.Value / 1000
        section(3).length = NumericUpDown56.Value / 1000
        section(3).k_stiffness = PI * g_modulus * section(3).dia ^ 4 / (32 * section(3).length)

        k_total = 1 / (1 / section(0).k_stiffness + 1 / section(1).k_stiffness + 1 / section(2).k_stiffness)

        If Not Double.IsNaN(k_total) And Not Double.IsInfinity(k_total) Then   'preventing NaN problem and infinity problems
            TextBox66.Text = Round(k_total, 0).ToString
            TextBox259.Text = Round(k_total, 0).ToString
            TextBox258.Text = Round(k_total * PI / 180, 0).ToString       'Drive shaft, Convert rad to degree
        End If
        'De as tussen de waaiers bij meerstrappers
        k2 = section(3).k_stiffness
        If Not Double.IsNaN(k2) And Not Double.IsInfinity(k2) Then   'preventing NaN problem and infinity problems
            TextBox220.Text = Round(k2, 0).ToString
        End If
    End Sub

    Private Sub Torsional_analyses()
        Dim omega, ii As Double

        Try
            Inertia_1 = NumericUpDown45.Value                       'Waaier #1 [kg.m2]
            Inertia_2 = NumericUpDown43.Value                       'Koppeling [kg.m2]
            Inertia_3 = NumericUpDown46.Value                       'Motor [kg.m2]
            Inertia_4 = NumericUpDown45.Value                       'Waaier #2 [kg.m2]

            Double.TryParse(TextBox259.Text, Springstiff_1)         'stijfheid as [Nm/rad]
            Springstiff_2 = NumericUpDown44.Value                   'stijfheid koppeling[Nm/rad]
            Springstiff_3 = section(3).k_stiffness

            For ii = 0 To 100
                omega = ii * NumericUpDown48.Value                              'Hoeksnelheid step-range
                Torsional_point(ii, 0) = Round(omega * 60 / (2 * PI), 0)        '[rad/s --> rpm]
                Torsional_point(ii, 1) = calc_zeroTorsion_4(omega)              'Residual torque
            Next

            draw_chart4()
            find_zero_torque()

        Catch ex As Exception
            MessageBox.Show("Problem torsional calculation")
        End Try
    End Sub

    Private Sub find_zero_torque()
        Dim T1, T2, T3, omg1, omg2, omg3 As Double
        Dim jj As Integer

        'TextBox158.Clear()

        omg1 = 1        'Start lower limit [rad/sec]
        omg2 = 300      'Start upper limit [rad/sec]
        omg3 = 3      'In the middle [rad/sec]

        T1 = calc_zeroTorsion_4(omg1)
        T2 = calc_zeroTorsion_4(omg2)
        T3 = calc_zeroTorsion_4(omg3)

        '-------------Iteratie 30x halveren moet voldoende zijn ---------------
        For jj = 0 To 30
            If T1 * T3 < 0 Then
                omg2 = omg3
            Else
                omg1 = omg3
            End If
            omg3 = (omg1 + omg2) / 2
            T1 = calc_zeroTorsion_4(omg1)
            T2 = calc_zeroTorsion_4(omg2)
            T3 = calc_zeroTorsion_4(omg3)
        Next
        TextBox84.Text = Round((omg3 * 60 / (2 * PI)), 0)        '[rad/s --> rpm]
        If T3 > 1 Then   'Residual torque too big,  problem in choosen bouderies
            TextBox84.BackColor = Color.Red
        Else
            TextBox84.BackColor = SystemColors.Window
        End If
    End Sub

    'Holzer residual torque analyses
    Private Function calc_zeroTorsion_4(omega As Double)
        Dim theta_1, theta_2, theta_3, theta_4 As Double
        Dim Torsion_1, Torsion_2, Torsion_3, Torsion_4 As Double

        theta_1 = 1                                             'Initial hoek verdraaiiing
        Torsion_1 = (omega ^ 2) * Inertia_1 * theta_1
        theta_2 = 1 - Torsion_1 / Springstiff_1                 'theta_1 - (((omega ^ 2) / Springstiff_1) * Inertia_1 * theta_1)
        Torsion_2 = Torsion_1 + (omega ^ 2) * Inertia_2 * theta_2
        theta_3 = theta_2 - Torsion_2 / Springstiff_2           'theta_2 - ((omega ^ 2) / Springstiff_2) * (Inertia_1 * theta_1 + Inertia_2 * theta_2)
        Torsion_3 = Torsion_2 + (omega ^ 2) * Inertia_3 * theta_3
        theta_4 = theta_3 - Torsion_3 / Springstiff_3
        Torsion_4 = Torsion_3 + (omega ^ 2) * Inertia_4 * theta_4

        If (RadioButton15.Checked) Then
            Return (Torsion_3)                 '[Nm] enkel trapper
        Else
            Return (Torsion_4)                 '[Nm] meer trapper
        End If

    End Function

    Private Sub draw_chart4()
        Dim hh As Integer
        Try
            Chart4.Series.Clear()
            Chart4.ChartAreas.Clear()
            Chart4.Titles.Clear()

            Chart4.Series.Add("Residual Torque")

            Chart4.ChartAreas.Add("ChartArea0")
            Chart4.Series(0).ChartArea = "ChartArea0"

            Chart4.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line

            Chart4.Titles.Add("Torsional natural frequency analysis")
            Chart4.Titles(0).Font = New Font("Arial", 16, System.Drawing.FontStyle.Bold)

            Chart4.Series(0).Name = "Torque"
            Chart4.Series(0).Color = Color.LightGreen
            Chart4.Series(0).BorderWidth = 1

            Chart4.ChartAreas("ChartArea0").AxisX.Title = "[rpm]"
            Chart4.ChartAreas("ChartArea0").AxisY.Title = "Torsion_3 [Nm] * 10^6"
            Chart4.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart4.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical
            Chart4.Series(0).YAxisType = AxisType.Primary

            For hh = 0 To 100
                Chart4.Series(0).Points.AddXY(Torsional_point(hh, 0), Torsional_point(hh, 1))
            Next
        Catch ex As Exception
            MessageBox.Show("nnnnnn")
        End Try
    End Sub

    Private Sub calc_bearing_belts()
        Dim length_a, length_b, length_c, length_naaf As Double
        Dim dia_a, dia_b, dia_c, dia_naaf As Double
        Dim g_shaft_a, g_shaft_b, g_shaft_c, sg_staal As Double
        Dim J_shaft_a, J_shaft_b, J_shaft_c, J_shaft_total As Double
        Dim I_shaft_a, I_shaft_b, I_shaft_c As Double
        Dim gewicht_as, gewicht_waaier, gewicht_pulley, gewicht_naaf As Double
        Dim Force_combi1, Force_combi2, Force_combi3 As Double
        Dim N_kritisch_as, N_max_tussen As Double
        Dim N_max_doorbuiging As Double
        Dim n_actual As Double
        Dim Elasticiteitsm As Double
        Dim motor_inertia As Double
        Dim W_rpm As Double
        Dim F_onbalans, V_onbalans, hoeksnelheid As Double
        Dim F_a_hor, F_a_vert, F_a_combined As Double       'Bearing next impellar
        Dim F_b_hor, F_b_vert, F_b_combined As Double       'Bearing next coupling
        Dim F_axial As Double
        Dim dia_pulley, S_power, F_snaar, F_scheef As Double
        Dim Waaier_dia, Voorplaat_keel As Double

        W_rpm = NumericUpDown19.Value                       'Toerental waaier
        TextBox110.Text = NumericUpDown19.Value.ToString

        length_a = NumericUpDown22.Value / 1000         '[m]
        length_b = NumericUpDown23.Value / 1000         '[m]
        length_c = NumericUpDown24.Value / 1000         '[m]
        length_naaf = NumericUpDown29.Value / 1000      '[m]

        dia_a = NumericUpDown25.Value / 1000            '[m]
        dia_b = NumericUpDown26.Value / 1000            '[m]
        dia_c = NumericUpDown27.Value / 1000            '[m]
        dia_naaf = NumericUpDown28.Value / 1000         '[m]

        Double.TryParse(TextBox33.Text, sg_staal)
        g_shaft_a = PI / 4 * dia_a ^ 2 * length_a * sg_staal
        g_shaft_b = PI / 4 * dia_b ^ 2 * length_b * sg_staal
        g_shaft_c = PI / 4 * dia_c ^ 2 * length_c * sg_staal
        gewicht_as = g_shaft_a + g_shaft_b + g_shaft_c

        Double.TryParse(TextBox374.Text, gewicht_waaier)
        gewicht_pulley = NumericUpDown30.Value
        gewicht_naaf = PI / 4 * dia_naaf ^ 2 * length_naaf * sg_staal

        I_shaft_a = PI * dia_a ^ 4 / 64                     'OppervlakTraagheid [m4]
        I_shaft_b = PI * dia_b ^ 4 / 64                     'OppervlakTraagheid [m4]
        I_shaft_c = PI * dia_c ^ 4 / 64                     'OppervlakTraagheid [m4]

        J_shaft_a = 0.5 * g_shaft_a * (dia_a / 2) ^ 2       'MassaTraagheid [kg.m2]
        J_shaft_b = 0.5 * g_shaft_b * (dia_b / 2) ^ 2       'MassaTraagheid [kg.m2]
        J_shaft_c = 0.5 * g_shaft_c * (dia_c / 2) ^ 2       'MassaTraagheid [kg.m2]
        J_shaft_total = J_shaft_a + J_shaft_b + J_shaft_c   'MassaTraagheid As


        '--Willi Bohl, Ventilatoren, Kritisch toerental formule 6.41 pagina 213--------------
        '----- waaier buiten de lagers ----------------------
        Elasticiteitsm = 210 * 1000 ^ 3                                             'in Pascal [N/m2]
        N_max_doorbuiging = ((length_a ^ 3 / I_shaft_a) + (length_a ^ 2 * length_b / I_shaft_b)) / (3 * Elasticiteitsm)
        N_max_doorbuiging *= (gewicht_waaier) * 9.81
        N_kritisch_as = Sqrt(9.81 / N_max_doorbuiging) * 60 / (2 * PI)

        '----- waaier tussen de lagers ----------------------
        N_max_tussen = Sqrt(48 * Elasticiteitsm * I_shaft_b / (gewicht_waaier * length_b)) * 60 / (2 * PI)


        '--------- Kracht door onbalans----------
        V_onbalans = NumericUpDown15.Value / 1000          '[m/s]
        hoeksnelheid = (2 * PI * W_rpm) / 60
        F_onbalans = V_onbalans * hoeksnelheid * (gewicht_waaier + gewicht_naaf + g_shaft_a)

        '--------- Kracht door V_snaren----------
        If (ComboBox4.SelectedIndex > -1) Then                                  'Prevent exceptions
            Dim words() As String = emotor_4P(ComboBox4.SelectedIndex).Split(";")
            S_power = words(0) * 1000                                           'Motor vermogen
            n_actual = words(1)                                                 'Toerental motor [rpm]
            dia_pulley = NumericUpDown16.Value / 1000
            F_snaar = 0.975 * S_power * 20 / (W_rpm * dia_pulley * 0.5)

            '------------- inertia motor--------------------
            motor_inertia = emotor_4P_inert(n_actual, S_power)
            TextBox219.Text = Round(motor_inertia, 1).ToString
            NumericUpDown46.Value = Round(motor_inertia, 1).ToString
        End If

        '--------- Scheefstelling koppeling ---------------
        F_scheef = 5.7 * Sqrt(S_power / W_rpm)                         '?????????????????????????????

        '----------- Forces bearing vertical-------------
        Force_combi1 = (gewicht_waaier + gewicht_naaf) * 9.81 + F_onbalans
        Force_combi2 = gewicht_as * 9.81
        Force_combi3 = gewicht_pulley * 9.81

        F_a_vert = Abs(Force_combi1 * (length_a + length_b) + Force_combi2 * length_b * 0.5 - Force_combi3 * length_c) / length_b
        F_b_vert = Abs(Force_combi1 * length_a - Force_combi2 * length_b * 0.5 - Force_combi3 * (length_b + length_c)) / length_b

        '----------- Forces bearing horizontal-------------
        Force_combi1 = 0   ' 
        Force_combi2 = 0   ' 
        If RadioButton7.Checked Then    'direct drive
            Force_combi3 = F_scheef
            F_snaar = 0
        Else
            Force_combi3 = F_snaar
            F_scheef = 0
        End If
        F_a_hor = Abs(Force_combi1 * (length_a + length_b) + Force_combi2 * length_b * 0.5 - Force_combi3 * length_c) / length_b
        F_b_hor = Abs(Force_combi1 * length_a - Force_combi2 * length_b * 0.5 - Force_combi3 * (length_b + length_c)) / length_b

        '----------- Forces bearing combined-------------
        F_a_combined = Sqrt(F_a_vert ^ 2 + F_a_hor ^ 2)
        F_b_combined = Sqrt(F_b_vert ^ 2 + F_b_hor ^ 2)

        '--------Axial force Keel diameter------------
        If ComboBox1.SelectedIndex > -1 Then '------- schoepgewicht berekenen-----------
            Waaier_dia = NumericUpDown21.Value / 1000 '[m]
            Voorplaat_keel = Tschets(ComboBox1.SelectedIndex).Tdata(16) / 1000 * (Waaier_dia / 1.0)     '[m]
            F_axial = PI / 4 * Voorplaat_keel ^ 2 * NumericUpDown37.Value * 100
        End If
        ' MessageBox.Show("Voorplaat_keel=" & Voorplaat_keel.ToString &"  F_b_hor =" & F_b_hor.ToString)

        '----------- Present massa traagheid-------------
        TextBox35.Text = Round(J_shaft_a, 2).ToString           'Massa traagheid (0.5*M*R^2)
        TextBox39.Text = Round(J_shaft_b, 2).ToString           'Massa traagheid (0.5*M*R^2)
        TextBox194.Text = Round(J_shaft_c, 2).ToString          'Massa traagheid (0.5*M*R^2)
        TextBox41.Text = Round(J_shaft_total, 2).ToString       'Massa traagheid (0.5*M*R^2)

        '----------- Present gewicht------------------
        TextBox46.Text = Round(g_shaft_a, 1).ToString
        TextBox48.Text = Round(g_shaft_b, 1).ToString
        TextBox52.Text = Round(g_shaft_c, 1).ToString
        TextBox189.Text = TextBox46.Text
        TextBox191.Text = TextBox48.Text
        TextBox193.Text = TextBox52.Text
        TextBox190.Text = Round(gewicht_as, 0).ToString
        TextBox102.Text = Round(g_shaft_a + g_shaft_b + g_shaft_c + gewicht_naaf + gewicht_pulley, 1).ToString 'Totaal gewicht impellar

        '--------------- krachten---------------
        TextBox97.Text = Round(F_onbalans, 0).ToString      'Force inbalans
        TextBox98.Text = Round(F_snaar, 0).ToString         'Force trekkracht snaar
        TextBox99.Text = Round(F_scheef, 0).ToString        'Force scheefstelling (geen snaar)
        TextBox100.Text = Round(F_a_combined, 0).ToString   'Force lager A hor+vert combined
        TextBox101.Text = Round(F_b_combined, 0).ToString   'Force lager B hor+vert combined
        TextBox19.Text = Round(F_axial, 0).ToString         'Force axial

        '----------- eigen frequentie ---------------------
        TextBox47.Text = Round(N_kritisch_as, 0).ToString               '[RPM] overhung
        TextBox111.Text = Round(N_max_doorbuiging * 1000, 3).ToString   'Max doorbuiging in [mm]
        TextBox377.Text = Round(N_max_tussen, 0).ToString               '[rpm] Tussen de lagers

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, NumericUpDown8.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, ComboBox1.SelectedIndexChanged, RadioButton4.CheckedChanged, RadioButton3.CheckedChanged, CheckBox4.CheckedChanged, NumericUpDown33.ValueChanged, ComboBox7.SelectedIndexChanged, RadioButton14.CheckedChanged, RadioButton13.CheckedChanged, RadioButton12.CheckedChanged, NumericUpDown58.ValueChanged, NumericUpDown37.ValueChanged, NumericUpDown76.ValueChanged, RadioButton6.CheckedChanged
        Update_selectie()
    End Sub
    Private Sub Update_selectie()
        NumericUpDown9.Value = NumericUpDown33.Value
        NumericUpDown21.Value = NumericUpDown33.Value       'Diameter waaier spanning berekening
        NumericUpDown10.Value = NumericUpDown13.Value
        If TabControl1.SelectedTab.Name = "TabPage1" Then
            ComboBox7.SelectedIndex = ComboBox1.SelectedIndex       'type selectie
        End If
        Scale_rules_applied(ComboBox1.SelectedIndex, NumericUpDown9.Value, NumericUpDown10.Value, NumericUpDown12.Value)
        Selectie_1()
    End Sub

    'VDI 3731 
    Function dL_oktaaf(laufz As Double, dia As Double, v_buiten As Double, okt As Double)
        Dim dl As Double
        Select Case True
            Case laufz > 0.06 And laufz <= 0.2
                dl = -1 * (6 + 12 * (Log10(okt * dia / v_buiten) - 0.13) ^ 2)
                TextBox376.BackColor = Color.LightGreen
            Case laufz > 0.2 And laufz < 0.56
                dl = -1 * (5 + 5 * (Log10(okt * dia / v_buiten) - 0.39) ^ 2)
                TextBox376.BackColor = Color.LightGreen
            Case Else
                TextBox376.BackColor = Color.Red
        End Select
        Return (dl)
    End Function
    'VDI 3731 OPEN Suction-Discharge gleigung 20
    Function dL_open_pipe(freq As Double, diameter As Double)
        Dim dl, RKH, v_geluid, radius As Double

        Double.TryParse(TextBox347.Text, v_geluid)  'geluidsnelheid
        radius = diameter / 2
        RKH = 2 * PI * freq * radius / v_geluid
        dl = 10 * Log10(2.3 * RKH ^ 2 / (1 + 2.3 * RKH ^ 2))
        Return (dl)
    End Function
    Private Function add_decibels(snd As Double())
        Dim Ltot As Double = 0
        Dim i As Integer

        For i = 0 To 8
            Ltot += 10 ^ (snd(i) / 10)
        Next
        'Ltot = 10 * Log10(10 ^ (snd(0) / 10) + 10 ^ (snd(1) / 10) + 10 ^ (snd(2) / 10) + 10 ^ (snd(3) / 10) + 10 ^ (snd(4) / 10) + 10 ^ (snd(5) / 10) + 10 ^ (snd(6) / 10) + 10 ^ (snd(7) / 10))
        Ltot = 10 * Log10(Ltot)

        Return (Ltot)
    End Function
    'Convert Sound power to pressure
    Private Function power_to_pressure(sound_power As Double)
        Dim distance As Double

        distance = 1    '[m]
        Return (sound_power - Abs(10 * Log10(1 / (4 * PI * distance ^ 2))))
    End Function
    'Calculate density
    '1= inlet, 2= outlet
    'Temperatures in celsius
    'Pressure is Barometric (static) absolute pressure in Pascal

    Private Function calc_density(density1 As Double, pressure1 As Double, pressure2 As Double, temperature1 As Double, temperature2 As Double)
        ' MessageBox.Show(density1.ToString & " " & pressure1.ToString & "  " & temperature1.ToString)
        'If (pressure1 < 80000) Or (pressure1 < 80000) Then
        '    MessageBox.Show("Density calculation warning Pressure must be in Pa absolute")
        'End If

        ' MessageBox.Show(" Ro1= " & density1.ToString & " P1= " & Round(pressure1, 0).ToString & " P2= " & Round(pressure2, 0).ToString & " T1= " & Round(temperature1, 1).ToString & " T2= " & Round(temperature2, 1).ToString)

        Return (density1 * (pressure2 / pressure1) * ((temperature1 + 273.15) / (temperature2 + 273.15)))
    End Function
    'Calculate the noise tab
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click, TabPage4.Enter, NumericUpDown7.ValueChanged, ComboBox9.SelectedIndexChanged, ComboBox10.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, CheckBox11.CheckedChanged, CheckBox12.CheckedChanged
        Calc_noise()
    End Sub

    'Calculate the noise tab
    Private Sub Calc_noise()
        Dim p_stat, P_tot, Act_flow_sec_noise, roww As Double
        Dim L_eff, n_imp, no_schoepen As Double
        Dim Kw, dia_fan_inlet, diameter_imp As Double
        Dim Suction_raw(9), Suction_damper(9), Suction_clean(9), Suction_dba(9) As Double           'Suction fan
        Dim Discharge_raw(9), Discharge_damper(9), Discharge_clean(9), discharge_dba(9) As Double   'Discharge fan
        Dim Discharge_pressure(9), Suction_pressure(9) As Double                                    'Sound pressure fan
        Dim keel_diameter As Double
        Dim casing_raw(9), casing_insulation(9), casing_clean(9), casing_dba(9) As Double           'Casing
        Dim hh As Integer
        Dim words() As String
        Dim DzDw, fan_size_factor As Double                                 'Diameter zuig/diameter waaier
        Dim Area_casing, casing_dikte, area_measure As Double               'Estimated Area casing one side!!
        Dim sound_speed As Double     '

        If (ComboBox1.SelectedIndex > -1) Then      'Prevent exceptions

            Label152.Text = "Waaier type " & Label1.Text
            Double.TryParse(TextBox159.Text, dia_fan_inlet)                 'Diameter suction
            p_stat = NumericUpDown37.Value * 100                            '[mBar]->[Pa] static 
            Double.TryParse(TextBox273.Text, P_tot)                         '[mBar]
            P_tot *= 100                                                    '[mBar]->[Pa]
            Double.TryParse(TextBox22.Text, Act_flow_sec_noise)             '[m3/hr]

            If RadioButton6.Checked Then
                Act_flow_sec_noise *= 2.0                                   'Double inlet fan type
            End If

            sound_speed = Sqrt(1.41 * NumericUpDown76.Value * 100 / NumericUpDown12.Value)   'K_lucht= 1.41
            Act_flow_sec_noise /= 3600                                      '[m3/sec]
            Double.TryParse(TextBox274.Text, Kw)                            'as vermogen
            Double.TryParse(TextBox74.Text, L_eff)                          'Efficiency [%]
            L_eff /= 100                                                    'Efficiency [-]
            roww = NumericUpDown12.Value                                    'Density [kg/Am3]
            n_imp = NumericUpDown13.Value                                   'toerental
            no_schoepen = Tschets(ComboBox1.SelectedIndex).Tdata(13)        'aantal schoepen.    
            dia_fan_inlet /= 1000                                           '[m]
            diameter_imp = NumericUpDown33.Value / 1000                     '[m] Diam impeller    
            DzDw = dia_fan_inlet / diameter_imp                             'Dia zuig / dia waaier

            fan_size_factor = diameter_imp * 1000 / Tschets(ComboBox1.SelectedIndex).Tdata(0)
            Area_casing = Tschets(ComboBox1.SelectedIndex).Tdata(20) * fan_size_factor * 7.43    'Fan oppervlak [m2]
            area_measure = (Sqrt(Area_casing) + 2) ^ 2                      'Meet oppervlak Hoogte+2m, Breed+2m
            TextBox343.Text = Round(area_measure, 2).ToString               'Meet oppervlak

            casing_dikte = NumericUpDown7.Value                             'Casing plaat dikte
            keel_diameter = Tschets(ComboBox1.SelectedIndex).Tdata(16) / 1000 * fan_size_factor  '[m]

            '------------------------ casing insulation--------------------       
            If (ComboBox9.SelectedIndex > -1) Then      'Prevent exceptions
                words = insulation_casing(ComboBox9.SelectedIndex).Split(";")
                For hh = 0 To 7
                    casing_insulation(hh) = words(hh + 2)
                Next
                TextBox293.Text = Round(casing_insulation(0), 0).ToString
                TextBox290.Text = Round(casing_insulation(1), 0).ToString
                TextBox292.Text = Round(casing_insulation(2), 0).ToString
                TextBox291.Text = Round(casing_insulation(3), 0).ToString
                TextBox289.Text = Round(casing_insulation(4), 0).ToString
                TextBox288.Text = Round(casing_insulation(5), 0).ToString
                TextBox286.Text = Round(casing_insulation(6), 0).ToString
                TextBox285.Text = Round(casing_insulation(7), 0).ToString
            End If

            '----------------------- inlet damper--------------------------
            If (ComboBox10.SelectedIndex > -1) Then      'Prevent exceptions
                words = inlet_damper(ComboBox10.SelectedIndex).Split(";")
                For hh = 0 To 7
                    Suction_damper(hh) = words(hh + 1)
                Next
                TextBox332.Text = Round(Suction_damper(0), 0).ToString
                TextBox331.Text = Round(Suction_damper(1), 0).ToString
                TextBox330.Text = Round(Suction_damper(2), 0).ToString
                TextBox329.Text = Round(Suction_damper(3), 0).ToString
                TextBox328.Text = Round(Suction_damper(4), 0).ToString
                TextBox327.Text = Round(Suction_damper(5), 0).ToString
                TextBox326.Text = Round(Suction_damper(6), 0).ToString
                TextBox325.Text = Round(Suction_damper(7), 0).ToString
            End If

            '----------------------- Discharge damper--------------------------
            If (ComboBox5.SelectedIndex > -1) Then      'Prevent exceptions
                words = inlet_damper(ComboBox5.SelectedIndex).Split(";")
                For hh = 0 To 7
                    Discharge_damper(hh) = words(hh + 1)
                Next
                TextBox351.Text = Round(Discharge_damper(0), 0).ToString
                TextBox246.Text = Round(Discharge_damper(1), 0).ToString
                TextBox245.Text = Round(Discharge_damper(2), 0).ToString
                TextBox244.Text = Round(Discharge_damper(3), 0).ToString
                TextBox243.Text = Round(Discharge_damper(4), 0).ToString
                TextBox242.Text = Round(Discharge_damper(5), 0).ToString
                TextBox241.Text = Round(Discharge_damper(6), 0).ToString
                TextBox240.Text = Round(Discharge_damper(7), 0).ToString
            End If


            '--------------- VDI 3731 Blatt 2, gleigung 7 + 9 --------------
            '-------------------- Lw4A--- Ausblaskanal (Sound Power)--------
            Dim lw4a, v_omtrek, L_spec_labour, L_laufzahl, Schaufel_hz As Double
            Dim dL_okt(8) As Double

            L_spec_labour = P_tot / roww                            'Spec arbeid [J/kg]
            L_laufzahl = n_imp / 60 * Sqrt(Act_flow_sec_noise / Pow(T_spec_labour / 9.81, 0.75)) / 157.8
            TextBox376.Text = Round(L_laufzahl, 3).ToString         'Laufzahl


            v_omtrek = diameter_imp * PI * n_imp / 60                           '[m/s]
            lw4a = 85.5 + 10 * Log10(p_stat * Act_flow_sec_noise * (1 / L_eff - 1)) + 27.7 * Log10(v_omtrek / sound_speed)

            dL_okt(0) = lw4a + dL_oktaaf(L_laufzahl, diameter_imp, v_omtrek, 63)
            dL_okt(1) = lw4a + dL_oktaaf(L_laufzahl, diameter_imp, v_omtrek, 125)
            dL_okt(2) = lw4a + dL_oktaaf(L_laufzahl, diameter_imp, v_omtrek, 250)
            dL_okt(3) = lw4a + dL_oktaaf(L_laufzahl, diameter_imp, v_omtrek, 500)
            dL_okt(4) = lw4a + dL_oktaaf(L_laufzahl, diameter_imp, v_omtrek, 1000)
            dL_okt(5) = lw4a + dL_oktaaf(L_laufzahl, diameter_imp, v_omtrek, 2000)
            dL_okt(6) = lw4a + dL_oktaaf(L_laufzahl, diameter_imp, v_omtrek, 4000)
            dL_okt(7) = lw4a + dL_oktaaf(L_laufzahl, diameter_imp, v_omtrek, 8000)

            '------------- Add Schaufelfrequenz-Pegelzuschlag---------------
            'Oktaaf aanvullen tot het gewenste overall-nivo is bereikt-----
            Schaufel_hz = no_schoepen * n_imp / 60          'Schaufelfrequenz

            Select Case Schaufel_hz
                Case Is <= 63 * 1.41
                    Do While (lw4a > add_decibels(dL_okt))
                        dL_okt(0) += 0.1
                    Loop
                Case Is <= 125 * 1.41
                    Do While (lw4a > add_decibels(dL_okt))
                        dL_okt(1) += 0.1
                    Loop
                Case Is <= 500 * 1.41
                    Do While (lw4a > add_decibels(dL_okt))
                        dL_okt(2) += 0.1
                    Loop
                Case Is <= 1000 * 1.41
                    Do While (lw4a > add_decibels(dL_okt))
                        dL_okt(3) += 0.1
                    Loop
                Case Is <= 2000 * 1.41
                    Do While (lw4a > add_decibels(dL_okt))
                        dL_okt(4) += 0.1
                    Loop
                Case Is <= 4000 * 1.41
                    Do While (lw4a > add_decibels(dL_okt))
                        dL_okt(5) += 0.1
                    Loop
            End Select

            Discharge_raw(0) = dL_okt(0)            'Lp63
            Discharge_raw(1) = dL_okt(1)            'Lp125
            Discharge_raw(2) = dL_okt(2)            'Lp250
            Discharge_raw(3) = dL_okt(3)            'LP500  
            Discharge_raw(4) = dL_okt(4)            'Lp1000
            Discharge_raw(5) = dL_okt(5)            'Lp2000  
            Discharge_raw(6) = dL_okt(6)            'Lp4000
            Discharge_raw(7) = dL_okt(7)            'Lp8000

            'Technischen Akustik, 2 auflage Springer-Verlag
            'Seite 220----------------
            Suction_raw(0) = Discharge_raw(0) - 2   'Lp63
            Suction_raw(1) = Discharge_raw(1) - 2   'Lp125
            Suction_raw(2) = Discharge_raw(2) - 2   'Lp250
            Suction_raw(3) = Discharge_raw(3) - 2   'LP500  
            Suction_raw(4) = Discharge_raw(4) - 2   'Lp1000
            Suction_raw(5) = Discharge_raw(5) - 2   'Lp2000  
            Suction_raw(6) = Discharge_raw(6) - 2   'Lp4000
            Suction_raw(7) = Discharge_raw(7) - 2   'Lp8000

            '------------------------------------------------------
            '------------------------------------------------------
            Dim Rv_plate_thickness, Lv_area, deltaLA As Double

            Lv_area = 4.343 * Log(Area_casing)                              '[m2] fan oppervlak
            deltaLA = 4.343 * Log(area_measure)                             'Loss trough Casing Wall
            Rv_plate_thickness = 17.71 + 5.86 * Log(casing_dikte)           '[mm] casing plaat dikte


            '----------- calc discharge clean-----------
            For i = 0 To (Discharge_clean.Length - 1)
                Discharge_clean(i) = Discharge_raw(i) - Discharge_damper(i)
            Next

            '----------- calc suction clean-----------
            For i = 0 To (Suction_clean.Length - 1)
                Suction_clean(i) = Suction_raw(i) - Suction_damper(i)
            Next

            '----------- Casing RAW-----------
            For i = 0 To (casing_raw.Length - 1)
                casing_raw(i) = Discharge_raw(i) - Rv_plate_thickness + Lv_area
            Next

            '----------- Casing Clean [dB]-----------
            For i = 0 To (casing_clean.Length - 1)
                casing_clean(i) = casing_raw(i) - casing_insulation(i) - deltaLA
            Next

            '--------------- Open discharge reduction-------------
            Dim Open_disch_red(9) As Double
            Dim Open_suction_red(9) As Double
            Dim uit_b, uit_h, dia_uit As Double


            Double.TryParse(TextBox160.Text, uit_b)                 'Uitlaat breedte [mm]
            Double.TryParse(TextBox161.Text, uit_h)                 'Uitlaat hoogte [mm]
            uit_b /= 1000   '[m]
            uit_h /= 1000   '[m]
            dia_uit = 2 * uit_b * uit_h / (uit_b + uit_h)           'Hydraulische diameter [m]

            Open_disch_red(0) = dL_open_pipe(63, dia_uit)
            Open_disch_red(1) = dL_open_pipe(125, dia_uit)
            Open_disch_red(2) = dL_open_pipe(250, dia_uit)
            Open_disch_red(3) = dL_open_pipe(500, dia_uit)
            Open_disch_red(4) = dL_open_pipe(1000, dia_uit)
            Open_disch_red(5) = dL_open_pipe(2000, dia_uit)
            Open_disch_red(6) = dL_open_pipe(4000, dia_uit)
            Open_disch_red(7) = dL_open_pipe(8000, dia_uit)

            If CheckBox11.Checked Then  'Open discharge 
                Label361.Text = "Free Discharge Power (SWL)"
                Panel3.Visible = True
                For i = 0 To (Discharge_clean.Length - 1)
                    Discharge_clean(i) += Open_disch_red(i)
                Next
            Else
                Panel3.Visible = False
                Label361.Text = "Remaining induct (SWL)"
            End If

            '--------------- Open Suction reduction-------------
            Double.TryParse(TextBox159.Text, dia_uit)                   'Inlet diameter
            dia_uit /= 1000                                             '[m]

            Open_suction_red(0) = dL_open_pipe(63, dia_uit)
            Open_suction_red(1) = dL_open_pipe(125, dia_uit)
            Open_suction_red(2) = dL_open_pipe(250, dia_uit)
            Open_suction_red(3) = dL_open_pipe(500, dia_uit)
            Open_suction_red(4) = dL_open_pipe(1000, dia_uit)
            Open_suction_red(5) = dL_open_pipe(2000, dia_uit)
            Open_suction_red(6) = dL_open_pipe(4000, dia_uit)
            Open_suction_red(7) = dL_open_pipe(8000, dia_uit)

            If CheckBox12.Checked Then  'Open suction 
                Label356.Text = "Free Suction Power (SWL)"
                Panel4.Visible = True
                For i = 0 To (Suction_clean.Length - 1)
                    Suction_clean(i) += Open_suction_red(i)
                Next
            Else
                Panel4.Visible = False
                Label356.Text = "Remaining induct (SWL)"
            End If


            '------------- Power to pressure @ 1m --------------
            For i = 0 To (Discharge_clean.Length - 1)
                Discharge_pressure(i) = power_to_pressure(Discharge_clean(i))
                Suction_pressure(i) = power_to_pressure(Suction_clean(i))
            Next


            TextBox347.Text = Round(sound_speed, 1).ToString            '[m/s] geluidsnelhied
            TextBox114.Text = Round(n_imp, 0).ToString                  'Toerental [rpm]
            TextBox234.Text = Round(Kw, 0).ToString                     'As vermogen[kW]
            TextBox235.Text = Round(L_eff, 2).ToString                  'Efficiency [-]
            TextBox236.Text = Round(no_schoepen, 1).ToString            'Aantal Schoepen

            TextBox321.Text = Round(dia_fan_inlet, 2).ToString          'Dia zuig
            TextBox237.Text = Round(diameter_imp, 2).ToString           'Dia waaier
            TextBox381.Text = Round(keel_diameter, 2).ToString          'Dia Keel
            TextBox322.Text = Round(Area_casing, 1).ToString            'Estimated area fan casing one side [m2]


            TextBox316.Text = Round(Lv_area, 1).ToString                'Casing area [m2] 
            TextBox315.Text = Round(deltaLA, 1).ToString                'Meet area [m2] 
            TextBox314.Text = Round(Rv_plate_thickness, 3).ToString     'Plaat dikte isolatie 

            '---------------- input data--------------------------
            TextBox112.Text = Round(Act_flow_sec_noise * 3600, 0).ToString  'Debiet [m3/hr]
            TextBox127.Text = Round(Act_flow_sec_noise, 2).ToString         'Debiet [m3/s]
            TextBox113.Text = Round(p_stat / 100, 0).ToString               'Dp static [mbar]
            TextBox126.Text = Round(p_stat, 0).ToString                     'Dp static [Pa]


            '----------Suction RAW induct opgesplits in banden--------------
            TextBox115.Text = Round(Suction_raw(0), 1).ToString       '
            TextBox116.Text = Round(Suction_raw(1), 1).ToString
            TextBox117.Text = Round(Suction_raw(2), 1).ToString
            TextBox118.Text = Round(Suction_raw(3), 1).ToString
            TextBox119.Text = Round(Suction_raw(4), 1).ToString
            TextBox120.Text = Round(Suction_raw(5), 1).ToString
            TextBox121.Text = Round(Suction_raw(6), 1).ToString
            TextBox122.Text = Round(Suction_raw(7), 1).ToString
            TextBox125.Text = Round(add_decibels(Suction_raw), 1).ToString

            '----------Suction Clean induct opgesplits in banden--------------
            TextBox341.Text = Round(Suction_clean(0), 1).ToString       '
            TextBox340.Text = Round(Suction_clean(1), 1).ToString
            TextBox339.Text = Round(Suction_clean(2), 1).ToString
            TextBox338.Text = Round(Suction_clean(3), 1).ToString
            TextBox337.Text = Round(Suction_clean(4), 1).ToString
            TextBox336.Text = Round(Suction_clean(5), 1).ToString
            TextBox335.Text = Round(Suction_clean(6), 1).ToString
            TextBox334.Text = Round(Suction_clean(7), 1).ToString
            TextBox333.Text = Round(add_decibels(Suction_clean), 1).ToString

            '----------Discharge RAW induct opgesplits in banden--------------
            TextBox311.Text = Round(Discharge_raw(0), 1).ToString       '
            TextBox310.Text = Round(Discharge_raw(1), 1).ToString
            TextBox309.Text = Round(Discharge_raw(2), 1).ToString
            TextBox308.Text = Round(Discharge_raw(3), 1).ToString
            TextBox307.Text = Round(Discharge_raw(4), 1).ToString
            TextBox306.Text = Round(Discharge_raw(5), 1).ToString
            TextBox305.Text = Round(Discharge_raw(6), 1).ToString
            TextBox304.Text = Round(Discharge_raw(7), 1).ToString
            TextBox303.Text = Round(add_decibels(Discharge_raw), 1).ToString

            '----------Discharge Clean induct opgesplits in banden--------------
            TextBox369.Text = Round(Discharge_clean(0), 1).ToString       '
            TextBox368.Text = Round(Discharge_clean(1), 1).ToString
            TextBox367.Text = Round(Discharge_clean(2), 1).ToString
            TextBox366.Text = Round(Discharge_clean(3), 1).ToString
            TextBox365.Text = Round(Discharge_clean(4), 1).ToString
            TextBox364.Text = Round(Discharge_clean(5), 1).ToString
            TextBox363.Text = Round(Discharge_clean(6), 1).ToString
            TextBox362.Text = Round(Discharge_clean(7), 1).ToString
            TextBox361.Text = Round(add_decibels(Discharge_clean), 1).ToString

            '---------------- Casing raw -----------------------
            TextBox233.Text = Round(casing_raw(0), 1).ToString
            TextBox232.Text = Round(casing_raw(1), 1).ToString
            TextBox231.Text = Round(casing_raw(2), 1).ToString
            TextBox230.Text = Round(casing_raw(3), 1).ToString
            TextBox229.Text = Round(casing_raw(4), 1).ToString
            TextBox228.Text = Round(casing_raw(5), 1).ToString
            TextBox227.Text = Round(casing_raw(6), 1).ToString
            TextBox226.Text = Round(casing_raw(7), 1).ToString
            TextBox225.Text = Round(add_decibels(casing_raw), 1).ToString

            '---------------- Casing clean [dB]-----------------------
            TextBox302.Text = Round(casing_clean(0), 1).ToString
            TextBox299.Text = Round(casing_clean(1), 1).ToString
            TextBox301.Text = Round(casing_clean(2), 1).ToString
            TextBox300.Text = Round(casing_clean(3), 1).ToString
            TextBox298.Text = Round(casing_clean(4), 1).ToString
            TextBox297.Text = Round(casing_clean(5), 1).ToString
            TextBox295.Text = Round(casing_clean(6), 1).ToString
            TextBox294.Text = Round(casing_clean(7), 1).ToString
            TextBox296.Text = Round(add_decibels(casing_clean), 1).ToString

            '----------Open Discharge PRESSURE opgesplits in banden--------------
            TextBox356.Text = Round(Discharge_pressure(0), 1).ToString
            TextBox353.Text = Round(Discharge_pressure(1), 1).ToString
            TextBox355.Text = Round(Discharge_pressure(2), 1).ToString
            TextBox354.Text = Round(Discharge_pressure(3), 1).ToString
            TextBox352.Text = Round(Discharge_pressure(4), 1).ToString
            TextBox350.Text = Round(Discharge_pressure(5), 1).ToString
            TextBox349.Text = Round(Discharge_pressure(6), 1).ToString
            TextBox348.Text = Round(Discharge_pressure(7), 1).ToString
            TextBox238.Text = Round(add_decibels(Discharge_pressure), 1).ToString


            '----------Open Suction PRESSURE opgesplits in banden--------------
            TextBox317.Text = Round(Suction_pressure(0), 1).ToString
            TextBox313.Text = Round(Suction_pressure(1), 1).ToString
            TextBox318.Text = Round(Suction_pressure(2), 1).ToString
            TextBox319.Text = Round(Suction_pressure(3), 1).ToString
            TextBox320.Text = Round(Suction_pressure(4), 1).ToString
            TextBox323.Text = Round(Suction_pressure(5), 1).ToString
            TextBox324.Text = Round(Suction_pressure(6), 1).ToString
            TextBox357.Text = Round(Suction_pressure(7), 1).ToString
            TextBox312.Text = Round(add_decibels(Suction_pressure), 1).ToString
        End If
    End Sub

    'Calculate density at Normal Conditions
    'Normaal condities; 0 celsius, 101325 Pascal
    Private Function calc_Normal_density(density1 As Double, pressure As Double, temperature As Double)
        Return (density1 * (101325 / pressure) * ((temperature + 273.15) / 273.15))
    End Function

    'Calculate density at Actual Conditions
    'Normaal condities; 0 celsius, 101325 Pascal
    Private Function calc_Actual_density(density1 As Double, pressure As Double, temperature As Double)
        Return (density1 * (pressure / 101325) * (273.15 / (temperature + 273.15)))
    End Function

    'Calculate the labyrinth loss
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click, NumericUpDown55.ValueChanged, NumericUpDown54.ValueChanged, NumericUpDown53.ValueChanged, NumericUpDown52.ValueChanged, NumericUpDown51.ValueChanged, NumericUpDown50.ValueChanged, TabPage9.Enter, NumericUpDown84.ValueChanged, NumericUpDown83.ValueChanged, NumericUpDown82.ValueChanged
        Dim as_diam, spalt_breed, spalt_opp, rho, dpres, spalt_velos, no_rings, spalt_loss, contractie As Double

        '---------- as afdichting loss ------------------
        as_diam = NumericUpDown51.Value                     '[mm]
        spalt_breed = NumericUpDown50.Value                 '[mm]
        rho = NumericUpDown53.Value                         '[kg/m3]
        no_rings = NumericUpDown55.Value                    '[-]
        contractie = NumericUpDown54.Value                  '[-]
        dpres = NumericUpDown52.Value * 100 / no_rings      '[Pa] pressure loss per ring

        spalt_opp = PI * as_diam * spalt_breed              '[mm2]

        'Principle pressure is transferred into speed
        spalt_velos = contractie * Sqrt(dpres * 2 / rho)

        spalt_loss = spalt_velos * spalt_opp / 1000 ^ 2 * 36000 * rho   '[kg/hr]

        TextBox143.Text = Round(spalt_opp, 0).ToString      '[mm2]
        TextBox144.Text = Round(spalt_velos, 1).ToString    '[m/s]
        TextBox145.Text = Round(spalt_loss, 1).ToString     '[kg/hr]

        '------------ Labyrinth loss -----------------
        Dim Laby_diam, Laby_promille, Laby_contr, Laby_opp, Laby_velos, Laby_loss As Double

        Laby_diam = NumericUpDown84.Value                   '[mm]
        Laby_promille = NumericUpDown82.Value               '[0/000]
        Laby_contr = NumericUpDown83.Value                  '[-]

        Laby_opp = PI * Laby_diam ^ 2 * Laby_promille / 1000    '[mm2]

        'Principle pressure is transferred into speed
        Laby_velos = contractie * Sqrt(dpres * 2 / rho)

        Laby_loss = Laby_velos * Laby_opp / 1000 ^ 2 * 36000 * rho   '[kg/hr]

        TextBox379.Text = Round(Laby_opp, 0).ToString       '[mm2]
        TextBox378.Text = Round(Laby_velos, 1).ToString     '[m/s]
        TextBox375.Text = Round(Laby_loss, 1).ToString      '[kg/hr]
    End Sub


    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click, CheckBox6.CheckedChanged, CheckBox3.CheckedChanged, CheckBox2.CheckedChanged, TabPage3.Enter, NumericUpDown9.ValueChanged, NumericUpDown10.ValueChanged, RadioButton9.CheckedChanged, RadioButton11.CheckedChanged, RadioButton10.CheckedChanged, CheckBox7.CheckedChanged, CheckBox8.CheckedChanged, CheckBox10.CheckedChanged, CheckBox13.CheckedChanged, CheckBox14.CheckedChanged
        NumericUpDown33.Value = NumericUpDown9.Value
        If NumericUpDown33.Value > 2300 Then
            NumericUpDown33.BackColor = Color.Red
        Else
            NumericUpDown33.BackColor = Color.Yellow
        End If
        NumericUpDown21.Value = NumericUpDown9.Value    'Diameter waaier
        If NumericUpDown9.Value > 2300 Then
            NumericUpDown9.BackColor = Color.Red
        Else
            NumericUpDown9.BackColor = Color.Yellow
        End If
        NumericUpDown13.Value = NumericUpDown10.Value
        Scale_rules_applied(ComboBox1.SelectedIndex, NumericUpDown9.Value, NumericUpDown10.Value, NumericUpDown12.Value)
        draw_chart1(ComboBox1.SelectedIndex)
    End Sub
    'Calculate a impellar stage process condition inlet and outlet
    ' Gebaseerd op de schaal regels voor ventilatoren

    Private Sub Calc_stage(ByRef y As Stage)
        Dim area_uitlaat_flens As Double

        y.in_velos = y.Q1 / (PI / 4 * (y.zuig_dia / 1000) ^ 2)                  'Zuigmond snelheid [m/s]

        y.uitlaat_h = Round(Tschets(y.Typ).Tdata(4) * y.Dia1 / y.Dia0, 0)       'Uitlaat hoogte inw.[mm]
        y.uitlaat_b = Round(Tschets(y.Typ).Tdata(5) * y.Dia1 / y.Dia0, 0)       'Uitlaat breedte inw.[mm]
        area_uitlaat_flens = y.uitlaat_b * y.uitlaat_h / 10 ^ 6                 'Oppervlak uitlaatflens [m2]
        y.uit_velos = y.Q1 / area_uitlaat_flens                                 'Snelheid uitlaat [m/s]


        y.Q1 = Scale_rule_cap(y.Q0, y.Dia0, y.Dia1, y.Rpm0, y.Rpm1)                                 '[Am3/s]
        y.Pt2 = y.Pt1 + Scale_rule_Pressure(y.Pt0, y.Dia0, y.Dia1, y.Rpm0, y.Rpm1, y.Ro0, y.Ro1)    '[Pa]
        y.Ps2 = y.Ps1 + Scale_rule_Pressure(y.Ps0, y.Dia0, y.Dia1, y.Rpm0, y.Rpm1, y.Ro0, y.Ro1)    '[Pa]
        y.Power = Scale_rule_Power(y.Power0, y.Dia0, y.Dia1, y.Rpm0, y.Rpm1, y.Ro0, y.Ro1)          '[kW]

        y.delta_pt = y.Pt2 - y.Pt1  'Drukverhoging waaier [Pa] total
        y.delta_ps = y.Ps2 - y.Ps1  'Drukverhoging waaier [Pa] static

        y.Eff = y.Q1 * y.Pt2 / (y.Power * 1000)                                     'Eff =Press*Volume/Power


        y.T2 = y.T1 + (y.Power * 3600 / (cp_air * y.Qkg))                           'Temperature outlet flange [celsius]
        y.Om_velos = PI * y.Dia1 / 1000 * y.Rpm1 / 60                               'Omtreksnelheid waaier


        T_reynolds = Round(T_omtrek_s * T_diaw_m / kin_visco_air(20), 0)            '---------- Renolds Tschets ----------------------------------
        y.Reynolds = Round(y.Om_velos * (y.Dia1 / 1000) / kin_visco_air(y.T1), 0)   '---------- Renolds actueel ---------------------------------

        y.Ackeret = 1 - 0.5 * (1.0 - Tschets(y.Typ).werkp_opT(0) / 100) * Pow((1 + (T_reynolds / y.Reynolds)), 0.2)

        y.Ro2 = calc_density(y.Ro1, (y.Ps1), (y.Ps2), y.T1, y.T2) 'Ro outlet flange fan
    End Sub

    Private Sub calc_loop_loss(ByRef x As Stage)
        Dim phi, area_omloop As Double

        '--------------------- gegevens omloop --------------------------
        x.uitlaat_h = Round(Tschets(x.Typ).Tdata(4) * x.Dia1 / x.Dia0, 0)       'Uitlaat hoogte inw.[mm]
        x.uitlaat_b = Round(Tschets(x.Typ).Tdata(5) * x.Dia1 / x.Dia0, 0)       'Uitlaat breedte inw.[mm]
        area_omloop = x.uitlaat_b * x.uitlaat_h / 10 ^ 6                        'Oppervlak omloop [m2]
        x.loop_velos = x.Q1 / area_omloop                                       'snelheid uitlaat [m/s]

        '----------------- actual drukverlies omloop  (3 bochten) -------
        phi = NumericUpDown58.Value
        x.loop_loss = 0.5 * phi * x.loop_velos ^ 2 * x.Ro1      '[Pa]
        x.loop_loss *= 3                                        '3x sharp bend

        '===== Druk verlies kan nooit groter zijn dan de begindruk =============
        If x.loop_loss > x.Pt2 Then
            x.loop_loss = 0.99 * x.Pt2
        End If
        If x.loop_loss > x.Ps2 Then
            x.loop_loss = 0.99 * x.Ps2
        End If

        x.Pt3 = x.Pt2 - x.loop_loss                                             '[Pa]
        x.Ps3 = x.Ps2 - x.loop_loss

        x.Ro3 = calc_density(x.Ro1, x.Ps1, x.Ps3, x.T1, x.T2) 'Static pressure
    End Sub

    Public Function Trend(Data() As PPOINT, ByVal Degree As Integer) As PPOINT()
        'degree 1 = straight line y=a+bx
        'degree n = polynomials!!

        Dim a(,), Ai(,), P(,) As Double             '2 Dimensional arrays
        Dim SigmaA(), SigmaP() As Double            '1 Dimensional arrays
        Dim PointCount, MaxTerm, m, n, i, j As Integer
        Dim Ret() As PPOINT
        Dim Equation As String

        Degree = Degree + 1

        MaxTerm = (2 * (Degree - 1))
        PointCount = Data.Length

        ReDim SigmaA(MaxTerm - 1)
        ReDim SigmaP(MaxTerm - 1)

        ' Get the coefficients lists for matrices A, and P
        For m = 0 To (MaxTerm - 1)
            For n = 0 To (PointCount - 1)
                ' MessageBox.Show(Data(n).x)
                SigmaA(m) = SigmaA(m) + (Data(n).x ^ (m + 1))
                SigmaP(m) = SigmaP(m) + ((Data(n).x ^ m) * Data(n).y)
            Next
        Next

        ' Create Matrix A, and fill in the coefficients
        ReDim a(Degree - 1, Degree - 1)

        For i = 0 To (Degree - 1)
            For j = 0 To (Degree - 1)
                If i = 0 And j = 0 Then
                    a(i, j) = PointCount
                Else
                    a(i, j) = SigmaA((i + j) - 1)
                End If
            Next
        Next

        ' Create Matrix P, and fill in the coefficients
        ReDim P(Degree - 1, 0)
        For i = 0 To (Degree - 1)
            P(i, 0) = SigmaP(i)
        Next

        ' We have A, and P of AB=P, so we can solve B because B=AiP
        Ai = MxInverse(a)
        BZ = MxMultiplyCV(Ai, P)

        ' Now we solve the equations and generate the list of points
        PointCount = PointCount - 1
        ReDim Ret(PointCount)

        ' Work out non exponential first term
        For i = 0 To PointCount
            Ret(i).x = Data(i).x
            Ret(i).y = BZ(0, 0)
        Next

        ' Work out other exponential terms including exp 1
        For i = 0 To PointCount
            For j = 1 To Degree - 1
                Ret(i).y = Ret(i).y + (BZ(j, 0) * Ret(i).x ^ j)
            Next
        Next

        '-------- show the coefficients-------------
        Equation = "y=" & Format$(BZ(0, 0), "0.00000") & " +"
        For j = 1 To Degree - 1
            Equation = Equation & Format$(BZ(j, 0), "0.00000") & "x^" & j & " +"
        Next
        Equation = Microsoft.VisualBasic.Left(Equation, Len(Equation) - 2)
        TextBox158.AppendText(Microsoft.VisualBasic.Left(Equation, Len(Equation) - 3) & Environment.NewLine)
        'MessageBox.Show(Equation)

        Trend = Ret
    End Function

    Public Function MxMultiplyCV(Matrix1(,) As Double, ColumnVector(,) As Double) As Double(,)

        Dim i, j As Integer
        Dim Rows, Cols As Integer
        Dim Ret(,) As Double        '2 Dimensional array

        Rows = Matrix1.GetLength(0) - 1
        Cols = Matrix1.GetLength(1) - 1

        ReDim Ret(ColumnVector.GetLength(0) - 1, 0) 'returns a column vector

        For i = 0 To Rows
            For j = 0 To Cols
                Ret(i, 0) = Ret(i, 0) + (Matrix1(i, j) * ColumnVector(j, 0))
            Next
        Next

        MxMultiplyCV = Ret
    End Function

    Public Function MxInverse(Matrix(,) As Double) As Double(,)
        Dim i, j As Integer
        Dim Rows, Cols As Integer       '1 Dimensional array
        Dim Tmp(,), Ret(,) As Double    '2 Dimensional array
        Dim Degree As Integer

        Tmp = Matrix

        Rows = Tmp.GetLength(0) - 1     'First dimension of the array
        Cols = Tmp.GetLength(1) - 1     'Second dimension of the array

        Degree = Cols + 1

        'Augment Identity matrix onto matrix M to get [M|I]

        ReDim Preserve Tmp(Rows, (Degree * 2) - 1)
        For i = Degree To (Degree * 2) - 1
            Tmp((i Mod Degree), i) = 1
        Next

        ' Now find the inverse using Gauss-Jordan Elimination which should get us [I|A-1]
        MxGaussJordan(Tmp)

        ' Copy the inverse (A-1) part to array to return
        ReDim Ret(Rows, Cols)
        For i = 0 To Rows
            For j = Degree To (Degree * 2) - 1
                Ret(i, j - Degree) = Tmp(i, j)
            Next
        Next

        MxInverse = Ret
    End Function
    ' https://social.msdn.microsoft.com/Forums/en-US/4b08ad2f-9ce7-4dd1-9e70-46d001549ffd/automating-microsoft-word-using-vbnet?forum=vblanguage
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara4 As Word.Paragraph

        'Start Word and open the document template. 
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering department"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = 16
        oPara1.Range.Font.Bold = True
        oPara1.Format.SpaceAfter = 2                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = 11
        oPara2.Format.SpaceAfter = 1
        oPara2.Range.Font.Bold = False
        oPara2.Range.Text = "This torsional analyses is (API-673) based on the Holzer method" & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 9
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        oTable.Cell(1, 1).Range.Text = "Project Name"
        oTable.Cell(1, 2).Range.Text = TextBox283.Text
        oTable.Cell(2, 1).Range.Text = "Project number "
        oTable.Cell(2, 2).Range.Text = "."
        oTable.Cell(3, 1).Range.Text = "Auther "
        oTable.Cell(3, 2).Range.Text = Environment.UserName
        oTable.Cell(4, 1).Range.Text = "Date "
        oTable.Cell(4, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(2)
        oTable.Rows.Item(1).Range.Font.Bold = True
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 14 x 5 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 18, 5)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 9
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        oTable.Cell(1, 1).Range.Text = "Equipment Data"
        oTable.Cell(1, 2).Range.Text = ""
        oTable.Cell(1, 3).Range.Text = ""

        oTable.Cell(2, 1).Range.Text = "Inertia impeller"
        oTable.Cell(2, 2).Range.Text = NumericUpDown45.Value
        oTable.Cell(2, 3).Range.Text = "[kg.m2]"

        oTable.Cell(3, 1).Range.Text = "Inertia coupling"
        oTable.Cell(3, 2).Range.Text = NumericUpDown43.Value
        oTable.Cell(3, 3).Range.Text = "[kg.m2]"

        oTable.Cell(4, 1).Range.Text = "Inertia motor"
        oTable.Cell(4, 2).Range.Text = NumericUpDown46.Value
        oTable.Cell(4, 3).Range.Text = "[kg.m2]"

        oTable.Cell(5, 1).Range.Text = "Shaft stiffness"
        oTable.Cell(5, 2).Range.Text = TextBox259.Text
        oTable.Cell(5, 3).Range.Text = "[N.m/rad]"
        oTable.Cell(5, 4).Range.Text = TextBox258.Text
        oTable.Cell(5, 5).Range.Text = "[k.N.m/°]"

        oTable.Cell(6, 1).Range.Text = "Coupling stiffness"
        oTable.Cell(6, 2).Range.Text = NumericUpDown44.Value
        oTable.Cell(6, 3).Range.Text = "[N.m/rad]"
        oTable.Cell(6, 4).Range.Text = TextBox257.Text
        oTable.Cell(6, 5).Range.Text = "[k.N.m/°]"

        '---- Shaft length-----
        oTable.Cell(7, 1).Range.Text = "Overhang impeller"
        oTable.Cell(7, 2).Range.Text = NumericUpDown22.Value
        oTable.Cell(7, 3).Range.Text = "[mm]"

        oTable.Cell(8, 1).Range.Text = "Distance between bearings"
        oTable.Cell(8, 2).Range.Text = NumericUpDown23.Value
        oTable.Cell(8, 3).Range.Text = "[mm]"

        oTable.Cell(9, 1).Range.Text = "Overhang coupling"
        oTable.Cell(9, 2).Range.Text = NumericUpDown24.Value
        oTable.Cell(9, 3).Range.Text = "[mm]"

        '---- Shaft diameter-----
        oTable.Cell(10, 1).Range.Text = "Diameter impeller shaft"
        oTable.Cell(10, 2).Range.Text = NumericUpDown25.Value
        oTable.Cell(10, 3).Range.Text = "[mm]"

        oTable.Cell(11, 1).Range.Text = "Diameter shaft bearings"
        oTable.Cell(11, 2).Range.Text = NumericUpDown26.Value
        oTable.Cell(11, 3).Range.Text = "[mm]"

        oTable.Cell(12, 1).Range.Text = "Diameter shaft coupling"
        oTable.Cell(12, 2).Range.Text = NumericUpDown27.Value
        oTable.Cell(12, 3).Range.Text = "[mm]"

        oTable.Cell(13, 1).Range.Text = "Weight impeller"
        oTable.Cell(13, 2).Range.Text = TextBox374.Text
        oTable.Cell(13, 3).Range.Text = "[kg]"


        '---- results torsie analyse---------
        oTable.Cell(14, 1).Range.Text = "Drive string 1st natural speed"
        oTable.Cell(14, 2).Range.Text = TextBox84.Text
        oTable.Cell(14, 3).Range.Text = "[rpm]"

        '---- results bending analyse---------
        oTable.Cell(15, 1).Range.Text = "Overhung Shaft 1st natural speed"
        oTable.Cell(15, 2).Range.Text = TextBox47.Text
        oTable.Cell(15, 3).Range.Text = "[rpm]"

        '---- results bending analyse---------
        oTable.Cell(16, 1).Range.Text = "Shaft between bearings 1st natural speed"
        oTable.Cell(16, 2).Range.Text = TextBox377.Text
        oTable.Cell(16, 3).Range.Text = "[rpm]"

        '---- results Impeller analyse---------
        oTable.Cell(17, 1).Range.Text = "Impeller back disk 1st natural speed"
        oTable.Cell(17, 2).Range.Text = TextBox211.Text
        oTable.Cell(17, 3).Range.Text = "[rpm]"
        oTable.Cell(17, 4).Range.Text = NumericUpDown17.Value
        oTable.Cell(17, 5).Range.Text = "[mm]"

        '---- results Impeller analyse---------
        oTable.Cell(18, 1).Range.Text = "Impeller front disk 1st natural speed"
        oTable.Cell(18, 2).Range.Text = TextBox371.Text
        oTable.Cell(18, 3).Range.Text = "[rpm]"
        oTable.Cell(18, 4).Range.Text = NumericUpDown31.Value
        oTable.Cell(18, 5).Range.Text = "[mm]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.7)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)
        oTable.Columns.Item(4).Width = oWord.InchesToPoints(0.6)
        oTable.Columns.Item(5).Width = oWord.InchesToPoints(0.8)

        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '------------------save picture ---------------- 
        Chart4.SaveImage("c:\Temp\MainChart.gif", System.Drawing.Imaging.ImageFormat.Gif)
        oPara4 = oDoc.Content.Paragraphs.Add
        oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oPara4.Range.InlineShapes.AddPicture("c:\Temp\MainChart.gif")
        oPara4.Range.InlineShapes.Item(1).LockAspectRatio = True
        'oPara4.Range.InlineShapes.Item(1).Width = 400
        oPara4.Range.InsertParagraphAfter()

    End Sub


    'Bearing calculation
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click, RadioButton24.CheckedChanged, RadioButton22.CheckedChanged, RadioButton19.CheckedChanged, RadioButton17.Click, NumericUpDown69.ValueChanged, TabPage13.Enter, NumericUpDown62.ValueChanged, NumericUpDown61.ValueChanged, ComboBox8.SelectedIndexChanged, NumericUpDown63.ValueChanged, NumericUpDown59.ValueChanged, RadioButton23.CheckedChanged, RadioButton18.CheckedChanged, NumericUpDown57.ValueChanged, RadioButton25.Click, RadioButton26.CheckedChanged, RadioButton27.CheckedChanged, RadioButton28.CheckedChanged
        Dim plain_bearing_area, plain_stress_no1, plain_stress_no2, plain_sliding_speed As Double
        Dim force_no1, force_no2, max_dyn_press, max_temp, max_speed, max_pv, actual_pv, n_rpm, shaft_dia As Double
        Dim visco, dia_at_center_ball As Double
        Dim a1_fact, aSKF_fact, p_fact, Exp_life, C_load, Equi_load As Double

        Try
            Select Case True                'Select bearing material
                Case RadioButton17.Checked  'Bronze bush
                    max_dyn_press = 25      'N/mm2
                    max_temp = 250          '[celsius]
                    max_speed = 0.5         '[m/s]
                    max_pv = 0              '[Mega.N.m2.m/s ==> N.mm2.m/s]
                Case RadioButton25.Checked  'Sintered bronze
                    max_dyn_press = 10
                    max_temp = 90
                    max_speed = 5
                    max_pv = 0
                Case RadioButton19.Checked  'Wrapped bronze
                    max_dyn_press = 40
                    max_temp = 150
                    max_speed = 1
                    max_pv = 0
                Case RadioButton22.Checked  'POM composite
                    max_dyn_press = 120
                    max_temp = 110
                    max_speed = 2.5
                    max_pv = 0
                Case RadioButton26.Checked  'PTFE composite
                    max_dyn_press = 80
                    max_temp = 250
                    max_speed = 2.0
                    max_pv = 0
                Case RadioButton24.Checked  'Filement wound
                    max_dyn_press = 140
                    max_temp = 140
                    max_speed = 0.5
                    max_pv = 0
                Case RadioButton27.Checked  'Bronze SAE 841
                    max_dyn_press = 14
                    max_temp = 104
                    max_speed = 6.5
                    max_pv = 1.75
                Case RadioButton28.Checked  'Vespel SP-21
                    max_dyn_press = 46.4
                    max_temp = 393
                    max_speed = 15.2
                    max_pv = 10.7
            End Select

            max_dyn_press = max_dyn_press * NumericUpDown69.Value   'Safety factor
            n_rpm = NumericUpDown63.Value                           '[rpm]
            shaft_dia = NumericUpDown61.Value

            'Get the force data from the other tab
            Double.TryParse(TextBox100.Text, force_no1)
            Double.TryParse(TextBox101.Text, force_no2)

            plain_bearing_area = shaft_dia * NumericUpDown62.Value
            plain_stress_no1 = Round(force_no1 / plain_bearing_area, 1)
            plain_stress_no2 = Round(force_no2 / plain_bearing_area, 1)

            '--------------------plain bearing sliding_speed ---------------------
            plain_sliding_speed = n_rpm / 60 * shaft_dia / 1000     '[m/s]
            actual_pv = plain_sliding_speed * plain_stress_no1

            '-------------------- bearing life ---------------------
            p_fact = 1
            If RadioButton18.Checked Then p_fact = 3        'Ball bearing
            If RadioButton23.Checked Then p_fact = 3.3      'Roller bearing

            Select Case ComboBox8.SelectedIndex
                Case 0
                    a1_fact = 1         '90%
                Case 1
                    a1_fact = 0.64      '95%
                Case 2
                    a1_fact = 0.55      '96%
                Case 3
                    a1_fact = 0.47      '97%
                Case 4
                    a1_fact = 0.37      '98%
                Case 5
                    a1_fact = 0.25      '99%
            End Select

            aSKF_fact = 1

            '------------- expected life---------
            C_load = NumericUpDown57.Value                             'Data from manufacturer [kN]
            Equi_load = force_no1 / 1000 * NumericUpDown59.Value       '[kN]
            Exp_life = a1_fact * aSKF_fact * (C_load / Equi_load) ^ p_fact
            Exp_life *= 10 ^ 6 / (n_rpm * 60)

            '------------- reference viscosity---------
            dia_at_center_ball = shaft_dia * 1.5                        'is niet exact moet zijn (d_binnen+d_buiten)/2
            If n_rpm < 100 Then
                visco = 4500 * n_rpm ^ -0.83 * dia_at_center_ball ^ -0.5
            Else
                visco = 4500 * n_rpm ^ -0.5 * dia_at_center_ball ^ -0.5
            End If

            '--------- presenting ------------------
            TextBox221.Text = force_no1.ToString
            TextBox222.Text = force_no2.ToString
            TextBox223.Text = "--"
            TextBox252.Text = max_speed.ToString
            TextBox253.Text = max_temp.ToString
            TextBox254.Text = max_dyn_press.ToString
            TextBox256.Text = max_pv.ToString                       'Max P.V #1 

            TextBox251.Text = Round(visco, 1).ToString              'Required visco
            TextBox249.Text = Round(Equi_load, 1).ToString          'Equivalent Load [k.N]
            TextBox248.Text = Round(Exp_life / 1000, 0).ToString    'Expected life [kilo.hr]
            TextBox247.Text = plain_bearing_area.ToString           'Plain bearing area
            TextBox224.Text = plain_stress_no1.ToString             'Actual load on Bearing #1
            TextBox239.Text = plain_stress_no2.ToString             'Actual load on Bearing #2
            TextBox255.Text = actual_pv.ToString                    'Actual P.V #1 

            TextBox250.Text = a1_fact.ToString

            '--------- red or green  ------------------
            If Exp_life < 50000 Then
                TextBox248.BackColor = Color.Red
            Else
                TextBox248.BackColor = Color.LightGreen
            End If

            If plain_sliding_speed > max_speed Then
                TextBox252.BackColor = Color.Red
            Else
                TextBox252.BackColor = Color.LightGreen
            End If

            If plain_stress_no1 > max_dyn_press Or plain_stress_no2 > max_dyn_press Then
                TextBox224.BackColor = Color.Red
                TextBox239.BackColor = Color.Red
            Else
                TextBox224.BackColor = Color.LightGreen
                TextBox239.BackColor = Color.LightGreen
            End If

        Catch ex As Exception
        End Try
    End Sub
    'Calculate Actula-> Normal conditions
    'http://www.engineeringtoolbox.com/scfm-acfm-icfm-d_1012.html
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click, NumericUpDown64.ValueChanged, NumericUpDown60.ValueChanged, NumericUpDown47.ValueChanged
        TextBox263.Text = Round(calc_Normal_density(NumericUpDown47.Value, NumericUpDown64.Value, NumericUpDown60.Value), 4) 'Normal Conditions
    End Sub
    'Calculate Normal-> Actual conditions
    'http://www.engineeringtoolbox.com/scfm-acfm-icfm-d_1012.html
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click, NumericUpDown71.ValueChanged, NumericUpDown70.ValueChanged, NumericUpDown68.ValueChanged
        TextBox264.Text = Round(calc_Actual_density(NumericUpDown68.Value, NumericUpDown71.Value, NumericUpDown70.Value), 4) 'Actual conditions
    End Sub
    'Pressure Units conversion
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click, NumericUpDown65.ValueChanged
        TextBox265.Text = Round(NumericUpDown65.Value * 9.80665, 0)           'Pascal
        TextBox266.Text = Round(NumericUpDown65.Value * 9.80665 / 100, 2)     'mbar
    End Sub

    Public Sub MxGaussJordan(Matrix(,) As Double)

        Dim Rows, Cols As Integer
        Dim P, i, j As Integer
        Dim m, d, Pivot As Double

        Rows = Matrix.GetLength(0) - 1     'First dimension of the array
        Cols = Matrix.GetLength(1) - 1     'Second dimension of the array

        ' Reduce so we get the leading diagonal
        For P = 0 To Rows
            Pivot = Matrix(P, P)
            For i = 0 To Rows
                If Not P = i Then
                    m = Matrix(i, P) / Pivot
                    For j = 0 To Cols
                        Matrix(i, j) = Matrix(i, j) + (Matrix(P, j) * -m)
                    Next
                End If
            Next
        Next

        'Divide through to get the identity matrix
        'Note: the identity matrix may have very small values (close to zero)
        'because of the way floating points are stored.
        For i = 0 To Rows
            d = Matrix(i, i)
            For j = 0 To Cols
                Matrix(i, j) = Matrix(i, j) / d
            Next
        Next
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click, TabPage14.Enter, CheckBox9.CheckedChanged, CheckBox5.CheckedChanged
        draw_chart5()   'Draw the cases
    End Sub
    'Save Case button is hit
    Private Sub Button18_Click_1(sender As Object, e As EventArgs) Handles Button18.Click

        '----------- Variable------------------
        case_x_conditions(0, 10) = "Case name"
        case_x_conditions(1, 10) = "Model"
        case_x_conditions(2, 10) = "Speed"
        case_x_conditions(3, 10) = "Diameter"
        case_x_conditions(4, 10) = "Mass flow"
        case_x_conditions(5, 10) = "Inlet diameter"
        case_x_conditions(6, 10) = "Outlet size"

        '----------- inlet data--------------------
        case_x_conditions(7, 10) = "Suction Flow"
        case_x_conditions(8, 10) = "Suction Flow"
        case_x_conditions(9, 10) = "Suction Temp."
        case_x_conditions(10, 10) = "Suction Pressure"
        case_x_conditions(11, 10) = "Suction Density"

        '----------- outlet data--------------------
        case_x_conditions(12, 10) = "Discharge Flow"
        case_x_conditions(13, 10) = "Discharge Temp."
        case_x_conditions(14, 10) = "Disch.Pres.static"
        case_x_conditions(15, 10) = "Discharge Density"

        '----------- performance-------------------
        case_x_conditions(16, 10) = "dP Static"
        case_x_conditions(17, 10) = "dP Dynamic"
        case_x_conditions(18, 10) = "dP Total"
        case_x_conditions(19, 10) = "Shaft power"
        case_x_conditions(20, 10) = "Efficiency"
        case_x_conditions(21, 10) = "Mol weight "

        '----------- Units------------------
        case_x_conditions(0, 11) = " "
        case_x_conditions(1, 11) = " "
        case_x_conditions(2, 11) = "[rpm]"
        case_x_conditions(3, 11) = "[mm]"
        case_x_conditions(4, 11) = "[kg/hr]"
        case_x_conditions(5, 11) = "[mm]"
        case_x_conditions(6, 11) = "[mm]"

        '----------- inlet data--------------------
        case_x_conditions(7, 11) = "[Am3/hr]"
        case_x_conditions(8, 11) = "[Nm3/hr]"
        case_x_conditions(9, 11) = "[°c]"
        case_x_conditions(10, 11) = "[mbar abs]"
        case_x_conditions(11, 11) = "[kg/Am3]"

        '----------- outlet data--------------------
        case_x_conditions(12, 11) = "[Am3/hr]"
        case_x_conditions(13, 11) = "[°c]"
        case_x_conditions(14, 11) = "[mbar abs]"
        case_x_conditions(15, 11) = "[kg/Am3]"

        '----------- performance-------------------
        case_x_conditions(16, 11) = "[mbar.g]"
        case_x_conditions(17, 11) = "[mbar.g]"
        case_x_conditions(18, 11) = "[mbar.g]"
        case_x_conditions(19, 11) = "[kW]"
        case_x_conditions(20, 11) = "[%]"
        case_x_conditions(21, 11) = "[g/mol]"

        '----------- general data------------------
        case_x_conditions(0, NumericUpDown72.Value) = TextBox89.Text                                'Case name 
        case_x_conditions(1, NumericUpDown72.Value) = Tschets(ComboBox1.SelectedIndex).Tname        'Model 
        case_x_conditions(2, NumericUpDown72.Value) = NumericUpDown13.Value.ToString                'Speed [rpm]
        case_x_conditions(3, NumericUpDown72.Value) = NumericUpDown33.Value.ToString                'Diameter [mm]
        case_x_conditions(4, NumericUpDown72.Value) = TextBox157.Text                               'Mass flow [kg/hr]
        case_x_conditions(5, NumericUpDown72.Value) = TextBox159.Text                               'Inlet Diameter [mm]
        case_x_conditions(6, NumericUpDown72.Value) = TextBox160.Text & "x" & TextBox161.Text       'Outlet diemsions [mm]

        '----------- inlet data--------------------
        case_x_conditions(7, NumericUpDown72.Value) = TextBox272.Text                   'Flow [Am3/hr]
        case_x_conditions(8, NumericUpDown72.Value) = TextBox269.Text                   'Flow [Nm3/hr]
        case_x_conditions(9, NumericUpDown72.Value) = NumericUpDown4.Value.ToString     'Temp [c]
        case_x_conditions(10, NumericUpDown72.Value) = NumericUpDown76.Value.ToString    'Pressure [mbar abs]
        case_x_conditions(11, NumericUpDown72.Value) = NumericUpDown12.Value.ToString 'Density [kg/Am3]

        '----------- outlet data--------------------
        case_x_conditions(12, NumericUpDown72.Value) = TextBox267.Text                  'Volume Flow [Am3/hr]
        case_x_conditions(13, NumericUpDown72.Value) = TextBox54.Text                   'Temp uit[c]
        case_x_conditions(14, NumericUpDown72.Value) = TextBox23.Text                   'Static Pressure [mbar abs]
        case_x_conditions(15, NumericUpDown72.Value) = TextBox268.Text                  'Density [kg/Am3]

        '----------- performance-------------------
        case_x_conditions(16, NumericUpDown72.Value) = TextBox271.Text                  'Static dP [mbar.g]
        case_x_conditions(17, NumericUpDown72.Value) = TextBox75.Text                   'Dynamic dP [mbar.g]
        case_x_conditions(18, NumericUpDown72.Value) = TextBox273.Text                  'Total dP [mbar.g]
        case_x_conditions(19, NumericUpDown72.Value) = TextBox274.Text                  'Shaft power [kW]
        case_x_conditions(20, NumericUpDown72.Value) = TextBox275.Text                  'Efficiency [%]
        If RadioButton3.Checked Then    'Density given or calculated
            case_x_conditions(21, NumericUpDown72.Value) = NumericUpDown8.Value.ToString
        Else
            case_x_conditions(21, NumericUpDown72.Value) = "n.a."
        End If


        Button11_Click(sender, New System.EventArgs())  'Draw chart1 (calculate the data points before storage)

        '----------- chart data -------------------
        For hh = 0 To 50
            case_x_flow(hh, NumericUpDown72.Value) = case_x_flow(hh, 0)
            case_x_Pstat(hh, NumericUpDown72.Value) = case_x_Pstat(hh, 0)
            case_x_Power(hh, NumericUpDown72.Value) = case_x_Power(hh, 0)
        Next hh

        Button17_Click(sender, New System.EventArgs())  'Draw chart5
    End Sub
    'Case number is changed
    Private Sub NumericUpDown72_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown72.ValueChanged
        TextBox89.Text = case_x_conditions(0, NumericUpDown72.Value)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click, NumericUpDown75.ValueChanged, NumericUpDown74.ValueChanged, NumericUpDown73.ValueChanged
        'Calculate the density with mol weight and Boltzmann constant
        'Ro= P * MW /(8.31432 * (T+273))
        TextBox270.Text = Round(NumericUpDown73.Value * NumericUpDown75.Value / (8.31432 * 1000 * (NumericUpDown74.Value + 273.15)), 5).ToString
    End Sub
    'Site altitude ambient prssure calculation ..
    'See https://en.wikipedia.org/wiki/Barometric_formula
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click, NumericUpDown6.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown77.ValueChanged
        Dim site_altitude, ambient_mbar, ambient_Gas_mol_weight, ambient_temp, temp As Double

        site_altitude = NumericUpDown1.Value
        ambient_Gas_mol_weight = NumericUpDown6.Value / 1000
        ambient_temp = NumericUpDown77.Value

        temp = -9.80665 * ambient_Gas_mol_weight * (site_altitude - 0) / (8.31432 * (ambient_temp + 273.15))
        ambient_mbar = 1013.25 * Pow(Math.E, temp)

        TextBox78.Text = Round(ambient_mbar, 2).ToString
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click, NumericUpDown80.ValueChanged, NumericUpDown42.ValueChanged
        Dim Pressure As Double
        Pressure = NumericUpDown80.Value - 20 * Log10(NumericUpDown42.Value) - 8
        TextBox346.Text = Round(Pressure, 2).ToString
    End Sub
    'Adding sound
    Private Sub Button21_Click_1(sender As Object, e As EventArgs) Handles Button21.Click, NumericUpDown79.ValueChanged, NumericUpDown78.ValueChanged
        Dim Total As Double
        Total = 10 * Log10(10 ^ (NumericUpDown78.Value / 10) + 10 ^ (NumericUpDown79.Value / 10))
        TextBox287.Text = Round(Total, 2).ToString
    End Sub

    Private Sub calc_emotor_4P()
        'see http://ecatalog.weg.net/files/wegnet/WEG-specification-of-electric-motors-50039409-manual-english.pdf
        'see http://electrical-engineering-portal.com/calculation-of-motor-starting-time-as-first-approximation
        Dim Ins_power, required_power, aanlooptijd, n_actual, rad, shaft_power As Double
        Dim m_torque_inrush, m_torque_max, m_torque_rated, m_torque_average As Double
        Dim impellar_inertia, motor_inertia, total_inertia As Double
        Dim ang_acceleration, C_acc, inertia_torque, fan_load_torque As Double

        '----------------- aanlooptijd---------------
        If (ComboBox6.SelectedIndex > -1) Then      'Prevent exceptions

            '--------- motor torque-------------
            Dim words() As String = emotor_4P(ComboBox6.SelectedIndex).Split(";")
            Ins_power = words(0) * 1000             'Geinstalleerd vermogen [Watt]

            Select Case True                        'Toerental motor [rpm] 50 Hz
                Case RadioButton30.Checked
                    n_actual = 3000
                Case RadioButton29.Checked
                    n_actual = 1500
                Case RadioButton31.Checked
                    n_actual = 1000
                Case RadioButton35.Checked
                    n_actual = 750

                Case RadioButton33.Checked           'Toerental motor [rpm] 60 Hz
                    n_actual = 3600
                Case RadioButton34.Checked
                    n_actual = 1800
                Case RadioButton32.Checked
                    n_actual = 1200
                Case RadioButton36.Checked
                    n_actual = 900
            End Select


            rad = n_actual / 60 * 2 * PI                    'Hoeksnelheid [rad/s]
            m_torque_rated = Ins_power / rad
            m_torque_inrush = m_torque_rated * NumericUpDown14.Value
            m_torque_max = m_torque_rated * NumericUpDown34.Value
            m_torque_average = 0.45 * (m_torque_inrush + m_torque_max)

            '---------- actual required fan power----------------
            required_power = NumericUpDown36.Value * Ins_power  '[kW]
            fan_load_torque = required_power / rad              '[N.m]

            '------------- inertia load--------------------
            Double.TryParse(TextBox109.Text, impellar_inertia)
            impellar_inertia = impellar_inertia * NumericUpDown35.Value ^ 2         'in case speed ratio impeller/motor 

            '------------- inertia motor--------------------
            motor_inertia = emotor_4P_inert(n_actual, Ins_power)

            total_inertia = impellar_inertia + motor_inertia    '[kg.m2]
            inertia_torque = total_inertia * ang_acceleration     '[N.m]

            '-------------- aanloop tijd---------------
            C_acc = 0.45 * (m_torque_inrush + m_torque_max) - (NumericUpDown38.Value * fan_load_torque)
            aanlooptijd = 2 * PI * n_actual * total_inertia / (60 * C_acc)
        End If

        TextBox195.Text = Round(n_actual, 0).ToString               'Toerental [rpm]
        TextBox196.Text = Round(rad, 0).ToString                    'Hoeksnelheid [rad/s]
        TextBox197.Text = Round(m_torque_inrush, 0).ToString        'Start torque [N.m]
        TextBox198.Text = Round(m_torque_max, 0).ToString           'Max torque [N.m]
        TextBox199.Text = Round(motor_inertia, 2).ToString          'Motor inertia [kg.m2]
        TextBox200.Text = Round(m_torque_rated, 0).ToString         'Rated torque [N.m]
        TextBox202.Text = Round(impellar_inertia, 1).ToString       'impellar inertia [kg.m2]
        TextBox207.Text = Round(motor_inertia, 1).ToString          'motor inertia [kg.m2]
        NumericUpDown46.Value = Round(motor_inertia, 1).ToString    'motor inertia [kg.m2]
        TextBox213.Text = Round(total_inertia, 1).ToString          'Total inertia[kg.m2]
        TextBox206.Text = Round(m_torque_average, 0).ToString       'Torque average [kg.m2]
        TextBox205.Text = Round(C_acc, 0).ToString                  'Effective acceleration torque [N.m]
        TextBox214.Text = Round(required_power / 1000, 0).ToString  'Fan power @ max speed [kw]
        TextBox215.Text = Round(fan_load_torque, 0).ToString        'Fan torque @ max speed [N.m]
        TextBox146.Text = Round(aanlooptijd, 1).ToString            'Aanlooptijd [s]

        '------- check geinstalleerd vermogen 15% safety --------------------
        Double.TryParse(TextBox274.Text, shaft_power)

        If (Ins_power < (shaft_power * 1000 * 1.15)) Then        '15% safety
            Label254.Visible = True
            ComboBox6.BackColor = Color.Red
        Else
            Label254.Visible = False
            ComboBox6.BackColor = Color.Yellow
        End If

        '------- check aanlooptijd --------------------
        If aanlooptijd > 45 Or aanlooptijd <= 0 Then
            TextBox146.BackColor = Color.Red
        Else
            TextBox146.BackColor = Color.LightGreen
        End If

        '------- check koppel --------------------
        If C_acc <= 0 Then
            TextBox205.BackColor = Color.Red
        Else
            TextBox205.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles TabPage12.Enter, NumericUpDown38.ValueChanged, NumericUpDown34.ValueChanged, NumericUpDown14.ValueChanged, ComboBox6.SelectedIndexChanged, NumericUpDown35.ValueChanged, NumericUpDown36.ValueChanged, RadioButton30.CheckedChanged, RadioButton36.CheckedChanged, RadioButton35.CheckedChanged, RadioButton34.CheckedChanged, RadioButton33.CheckedChanged, RadioButton32.CheckedChanged, RadioButton31.CheckedChanged, RadioButton29.CheckedChanged
        Calc_stress_impeller()
        calc_emotor_4P()
        draw_chart3()
    End Sub
    Private Sub draw_chart3()
        Dim hh As Integer
        Dim words() As String

        Try
            'Clear all series And chart areas so we can re-add them
            Chart3.Series.Clear()
            Chart3.ChartAreas.Clear()
            Chart3.Titles.Clear()
            Chart3.Series.Add("Series0")        'Koppel
            Chart3.ChartAreas.Add("ChartArea0")
            Chart3.Series(0).ChartArea = "ChartArea0"
            Chart3.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart3.Titles.Add("VSD motor koppel")
            Chart3.Titles(0).Font = New Font("Arial", 16, System.Drawing.FontStyle.Bold)
            Chart3.Series(0).Name = "Koppel motor [%]"
            Chart3.Series(0).Color = Color.Blue
            Chart3.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart3.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart3.ChartAreas("ChartArea0").AxisY.Title = "Koppel [-]"
            Chart3.ChartAreas("ChartArea0").AxisX.Title = "[Hz]"

            '------------------- Draw the lines in the chart--------------------
            For hh = 0 To (EXD_VSD_torque.Length - 1)
                words = EXD_VSD_torque(hh).Split(";")

                Chart3.Series(0).Points.AddXY(words(0), words(2))
            Next hh
            Chart3.Refresh()
        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Line 2400")  ' Show the exception's message.
        End Try

    End Sub
    ' see http://ecatalog.weg.net/files/wegnet/WEG-specification-of-electric-motors-50039409-manual-english.pdf
    Function emotor_4P_inert(rpm As Double, kw As Double)
        Dim motor_inertia As Double
        Select Case True
            Case rpm = 3000 Or rpm = 3600
                motor_inertia = 0.04 * (kw / 1000) ^ 0.9 * 1 ^ 2.5    '2 poles (1 pair) (3000 rpm) [kg.m2]
            Case rpm = 1500 Or rpm = 1800
                motor_inertia = 0.04 * (kw / 1000) ^ 0.9 * 2 ^ 2.5    '4 poles (2 pair) (1500 rpm) [kg.m2]
            Case rpm = 1000 Or rpm = 1200
                motor_inertia = 0.04 * (kw / 1000) ^ 0.9 * 3 ^ 2.5    '6 poles (3 pair) (1000 rpm) [kg.m2]
            Case rpm = 750 Or rpm = 900
                motor_inertia = 0.04 * (kw / 1000) ^ 0.9 * 4 ^ 2.5    '8 poles (4 pair) (750 rpm) [kg.m2]
            Case Else
                MessageBox.Show("Error occured in Motor Inertia calculation ")
        End Select
        Return (motor_inertia)
    End Function
    'Different Cases in one chart

    Private Sub draw_chart5()
        Dim hh, case_j As Integer
        Dim case_counter As Integer = 0
        Dim debiet As Double

        Chart5.Series.Clear()
        Chart5.Titles.Clear()
        Chart5.ChartAreas.Clear()
        Chart5.ChartAreas.Add("ChartArea0")



        Try
            For hh = 1 To 9 'Determine how many cases there are
                If Not String.IsNullOrEmpty(case_x_conditions(0, hh)) Then case_counter += 1
            Next

            '---------- Pressure cases  ------------------
            For hh = 1 To (case_counter)
                Chart5.Series.Add("Pressure " & case_x_conditions(0, hh))
                Chart5.Series(hh - 1).ChartArea = "ChartArea0"
                Chart5.Series(hh - 1).SmartLabelStyle.Enabled = True
                If CheckBox9.Checked Then
                    Chart5.Series(hh - 1).IsVisibleInLegend = True
                Else
                    Chart5.Series(hh - 1).IsVisibleInLegend = False
                End If
                Chart5.Series(hh - 1).BorderWidth = 1
            Next


            '---------- Power cases ------------------
            For hh = 1 To (case_counter)
                Chart5.Series.Add("Power " & case_x_conditions(0, hh))
                Chart5.Series(hh - 1 + case_counter).ChartArea = "ChartArea0"
                If CheckBox9.Checked Then
                    Chart5.Series(hh - 1 + case_counter).IsVisibleInLegend = True
                Else
                    Chart5.Series(hh - 1 + case_counter).IsVisibleInLegend = False
                End If
                Chart5.Series(hh - 1 + case_counter).BorderWidth = 1
            Next

            If CheckBox5.Checked Then
                Chart5.Titles.Add("Static Pressure and power")
            Else
                Chart5.Titles.Add("Static Pressure")
            End If
            Chart5.Titles(0).Font = New Font("Arial", 16, System.Drawing.FontStyle.Bold)

            Chart5.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart5.ChartAreas("ChartArea0").AxisX.Title = "Suction Volume flow [Am3/hr]"
            Chart5.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart5.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True

            Chart5.ChartAreas("ChartArea0").AxisY.Title = "Delta Static Pressure [mBar]"
            Chart5.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical
            Chart5.ChartAreas("ChartArea0").AxisY2.Enabled = AxisEnabled.True
            Chart5.ChartAreas("ChartArea0").AxisY2.Title = "Shaft power [kW]"
            Chart5.ChartAreas("ChartArea0").AxisY2.MinorTickMark.Enabled = True

            '----------------------------- pressure-----------------------
            For case_j = 1 To (case_counter)                  'Plot the cases 1...8
                Chart5.Series(case_j - 1).ChartType = DataVisualization.Charting.SeriesChartType.Line
                For hh = 0 To 50
                    debiet = case_x_flow(hh, case_j) * 3600 '/hr
                    Chart5.Series(case_j - 1).Points.AddXY(debiet, Round(case_x_Pstat(hh, case_j), 1))
                Next hh
                Chart5.Series(case_j - 1).Points(45).Label = "Pstat " & case_x_conditions(0, case_j)       'Case name 
            Next case_j

            '----------------------------- power-----------------------
            If CheckBox5.Checked Then
                For case_j = 1 To (case_counter)                 'Plot the cases 1...8
                    Chart5.Series(case_j - 1 + case_counter).ChartType = DataVisualization.Charting.SeriesChartType.Line
                    For hh = 0 To 50
                        debiet = case_x_flow(hh, case_j) * 3600 '/hr
                        Chart5.Series(case_j - 1 + case_counter).Points.AddXY(debiet, Round(case_x_Power(hh, case_j), 1))
                    Next hh
                    Chart5.Series(case_j - 1 + case_counter).Points(45).Label = "Power " & case_x_conditions(0, case_j)       'Case name 
                Next case_j
            End If

            Chart5.Refresh()
        Catch ex As Exception
            'MessageBox.Show(ex.Message & "Line 3332")  ' Show the exception's message.
        End Try

    End Sub

End Class
