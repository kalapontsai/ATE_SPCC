CREATE TABLE LotTitle (
lotdt_idx bigint,
lotname varchar(255) NOT NULL,
device varchar(255) NOT NULL,
tester int DEFAULT 0,
col_1_h float DEFAULT 0,
col_1_l float DEFAULT 0,
col_2_h float DEFAULT 0,
col_2_l float DEFAULT 0,
col_3_h float DEFAULT 0,
col_3_l float DEFAULT 0,
col_4_h float DEFAULT 0,
col_4_l float DEFAULT 0,
col_5_h float DEFAULT 0,
col_5_l float DEFAULT 0,
col_6_h float DEFAULT 0,
col_6_l float DEFAULT 0,
col_7_h float DEFAULT 0,
col_7_l float DEFAULT 0,
col_8_h float DEFAULT 0,
col_8_l float DEFAULT 0,
col_9_h float DEFAULT 0,
col_9_l float DEFAULT 0,
col_10_h float DEFAULT 0,
col_10_l float DEFAULT 0,
col_11_h float DEFAULT 0,
col_11_l float DEFAULT 0,
col_12_h float DEFAULT 0,
col_12_l float DEFAULT 0,
col_13_h float DEFAULT 0,
col_13_l float DEFAULT 0,
col_14_h float DEFAULT 0,
col_14_l float DEFAULT 0,
col_15_h float DEFAULT 0,
col_15_l float DEFAULT 0,
col_16_h float DEFAULT 0,
col_16_l float DEFAULT 0,
col_17_h float DEFAULT 0,
col_17_l float DEFAULT 0,
col_18_h float DEFAULT 0,
col_18_l float DEFAULT 0,
col_19_h float DEFAULT 0,
col_19_l float DEFAULT 0,
col_20_h float DEFAULT 0,
col_20_l float DEFAULT 0,
col_21_h float DEFAULT 0,
col_21_l float DEFAULT 0,
col_22_h float DEFAULT 0,
col_22_l float DEFAULT 0,
col_23_h float DEFAULT 0,
col_23_l float DEFAULT 0,
col_24_h float DEFAULT 0,
col_24_l float DEFAULT 0,
col_25_h float DEFAULT 0,
col_25_l float DEFAULT 0,
col_26_h float DEFAULT 0,
col_26_l float DEFAULT 0,
col_27_h float DEFAULT 0,
col_27_l float DEFAULT 0,
col_28_h float DEFAULT 0,
col_28_l float DEFAULT 0,
col_29_h float DEFAULT 0,
col_29_l float DEFAULT 0,
col_30_h float DEFAULT 0,
col_30_l float DEFAULT 0,
col_31_h float DEFAULT 0,
col_31_l float DEFAULT 0,
col_32_h float DEFAULT 0,
col_32_l float DEFAULT 0,
col_33_h float DEFAULT 0,
col_33_l float DEFAULT 0,
col_34_h float DEFAULT 0,
col_34_l float DEFAULT 0,
col_35_h float DEFAULT 0,
col_35_l float DEFAULT 0,
col_36_h float DEFAULT 0,
col_36_l float DEFAULT 0,
col_37_h float DEFAULT 0,
col_37_l float DEFAULT 0,
col_38_h float DEFAULT 0,
col_38_l float DEFAULT 0,
col_39_h float DEFAULT 0,
col_39_l float DEFAULT 0,
col_40_h float DEFAULT 0,
col_40_l float DEFAULT 0,
col_41_h float DEFAULT 0,
col_41_l float DEFAULT 0,
col_42_h float DEFAULT 0,
col_42_l float DEFAULT 0,
col_43_h float DEFAULT 0,
col_43_l float DEFAULT 0,
col_44_h float DEFAULT 0,
col_44_l float DEFAULT 0,
col_45_h float DEFAULT 0,
col_45_l float DEFAULT 0,
col_46_h float DEFAULT 0,
col_46_l float DEFAULT 0,
col_47_h float DEFAULT 0,
col_47_l float DEFAULT 0,
col_48_h float DEFAULT 0,
col_48_l float DEFAULT 0,
col_49_h float DEFAULT 0,
col_49_l float DEFAULT 0,
col_50_h float DEFAULT 0,
col_50_l float DEFAULT 0,
col_51_h float DEFAULT 0,
col_51_l float DEFAULT 0,
col_52_h float DEFAULT 0,
col_52_l float DEFAULT 0,
col_53_h float DEFAULT 0,
col_53_l float DEFAULT 0,
col_54_h float DEFAULT 0,
col_54_l float DEFAULT 0,
col_55_h float DEFAULT 0,
col_55_l float DEFAULT 0,
col_56_h float DEFAULT 0,
col_56_l float DEFAULT 0,
col_57_h float DEFAULT 0,
col_57_l float DEFAULT 0,
col_58_h float DEFAULT 0,
col_58_l float DEFAULT 0,
col_59_h float DEFAULT 0,
col_59_l float DEFAULT 0,
col_60_h float DEFAULT 0,
col_60_l float DEFAULT 0,
col_61_h float DEFAULT 0,
col_61_l float DEFAULT 0,
col_62_h float DEFAULT 0,
col_62_l float DEFAULT 0,
col_63_h float DEFAULT 0,
col_63_l float DEFAULT 0,
col_64_h float DEFAULT 0,
col_64_l float DEFAULT 0,
col_65_h float DEFAULT 0,
col_65_l float DEFAULT 0,
col_66_h float DEFAULT 0,
col_66_l float DEFAULT 0,
col_67_h float DEFAULT 0,
col_67_l float DEFAULT 0,
col_68_h float DEFAULT 0,
col_68_l float DEFAULT 0,
col_69_h float DEFAULT 0,
col_69_l float DEFAULT 0,
col_70_h float DEFAULT 0,
col_70_l float DEFAULT 0,
col_71_h float DEFAULT 0,
col_71_l float DEFAULT 0,
col_72_h float DEFAULT 0,
col_72_l float DEFAULT 0,
col_73_h float DEFAULT 0,
col_73_l float DEFAULT 0,
col_74_h float DEFAULT 0,
col_74_l float DEFAULT 0,
col_75_h float DEFAULT 0,
col_75_l float DEFAULT 0,
PRIMARY KEY (lotdt_idx)
);

CREATE TABLE TestResult (
    id int,
    Descript varchar(255) NOT NULL,
    PRIMARY KEY (id)
);

CREATE TABLE TestUnit (
    col int,
    unit varchar(255) NOT NULL,
    name varchar(255) NOT NULL,
    PRIMARY KEY (col)
);

CREATE TABLE TestData 
(idx int IDENTITY(1,1), 
lotdt_idx bigint NOT NULL,
FOREIGN KEY (lotdt_idx) REFERENCES LotTitle(lotdt_idx), 
t_result int NOT NULL,
FOREIGN KEY (t_result) REFERENCES TestResult(id),
col_1 float DEFAULT 0,
col_2 float DEFAULT 0,
col_3 float DEFAULT 0,
col_4 float DEFAULT 0,
col_5 float DEFAULT 0,
col_6 float DEFAULT 0,
col_7 float DEFAULT 0,
col_8 float DEFAULT 0,
col_9 float DEFAULT 0,
col_10 float DEFAULT 0,
col_11 float DEFAULT 0,
col_12 float DEFAULT 0,
col_13 float DEFAULT 0,
col_14 float DEFAULT 0,
col_15 float DEFAULT 0,
col_16 float DEFAULT 0,
col_17 float DEFAULT 0,
col_18 float DEFAULT 0,
col_19 float DEFAULT 0,
col_20 float DEFAULT 0,
col_21 float DEFAULT 0,
col_22 float DEFAULT 0,
col_23 float DEFAULT 0,
col_24 float DEFAULT 0,
col_25 float DEFAULT 0,
col_26 float DEFAULT 0,
col_27 float DEFAULT 0,
col_28 float DEFAULT 0,
col_29 float DEFAULT 0,
col_30 float DEFAULT 0,
col_31 float DEFAULT 0,
col_32 float DEFAULT 0,
col_33 float DEFAULT 0,
col_34 float DEFAULT 0,
col_35 float DEFAULT 0,
col_36 float DEFAULT 0,
col_37 float DEFAULT 0,
col_38 float DEFAULT 0,
col_39 float DEFAULT 0,
col_40 float DEFAULT 0,
col_41 float DEFAULT 0,
col_42 float DEFAULT 0,
col_43 float DEFAULT 0,
col_44 float DEFAULT 0,
col_45 float DEFAULT 0,
col_46 float DEFAULT 0,
col_47 float DEFAULT 0,
col_48 float DEFAULT 0,
col_49 float DEFAULT 0,
col_50 float DEFAULT 0,
col_51 float DEFAULT 0,
col_52 float DEFAULT 0,
col_53 float DEFAULT 0,
col_54 float DEFAULT 0,
col_55 float DEFAULT 0,
col_56 float DEFAULT 0,
col_57 float DEFAULT 0,
col_58 float DEFAULT 0,
col_59 float DEFAULT 0,
col_60 float DEFAULT 0,
col_61 float DEFAULT 0,
col_62 float DEFAULT 0,
col_63 float DEFAULT 0,
col_64 float DEFAULT 0,
col_65 float DEFAULT 0,
col_66 float DEFAULT 0,
col_67 float DEFAULT 0,
col_68 float DEFAULT 0,
col_69 float DEFAULT 0,
col_70 float DEFAULT 0,
col_71 float DEFAULT 0,
col_72 float DEFAULT 0,
col_73 float DEFAULT 0,
col_74 float DEFAULT 0,
col_75 float DEFAULT 0,
PRIMARY KEY (idx)
);

INSERT INTO TestUnit (col, unit, name)
VALUES
(1,'ma','空載輸入電流_L'),
(2,'ma','空載輸入電流_N'),
(3,'ma','空載輸入電流_H'),
(4,'W','空載輸入功率_L'),
(5,'W','空載輸入功率_N'),
(6,'W','空載輸入功率_H'),
(7,'V','空載輸出電壓CH1_L'),
(8,'V','空載輸出電壓CH1_N'),
(9,'V','空載輸出電壓CH1_H'),
(10,'V','空載輸出電壓CH2_L'),
(11,'V','空載輸出電壓CH2_N'),
(12,'V','空載輸出電壓CH2_H'),
(13,'p-pmV','空載漣波雜訊CH1_L'),
(14,'p-pmV','空載漣波雜訊CH1_N'),
(15,'p-pmV','空載漣波雜訊CH1_H'),
(16,'p-pmV','空載漣波雜訊CH2_L'),
(17,'p-pmV','空載漣波雜訊CH2_N'),
(18,'p-pmV','空載漣波雜訊CH2_H'),
(19,'mA','滿載輸入電流_L'),
(20,'mA','滿載輸入電流_N'),
(21,'mA','滿載輸入電流_H'),
(22,'W','滿載輸入功率_L'),
(23,'W','滿載輸入功率_N'),
(24,'W','滿載輸入功率_H'),
(25,'V','滿載輸出電壓CH1_L'),
(26,'V','滿載輸出電壓CH1_N'),
(27,'V','滿載輸出電壓CH1_H'),
(28,'V','滿載輸出電壓CH2_L'),
(29,'V','滿載輸出電壓CH2_N'),
(30,'V','滿載輸出電壓CH2_H'),
(31,'p-pmV','滿載漣波雜訊CH1_L'),
(32,'p-pmV','滿載漣波雜訊CH1_N'),
(33,'p-pmV','滿載漣波雜訊CH1_H'),
(34,'p-pmV','滿載漣波雜訊CH2_L'),
(35,'p-pmV','滿載漣波雜訊CH2_N'),
(36,'p-pmV','滿載漣波雜訊CH2_H'),
(37,'%','滿載效率_L'),
(38,'%','滿載效率_N'),
(39,'%','滿載效率_H'),
(40,'%','輸入電壓變動穩壓率CH1'),
(41,'%','輸入電壓變動穩壓率CH2'),
(42,'%','輸出負載變動穩壓率CH1_L'),
(43,'%','輸出負載變動穩壓率CH1_N'),
(44,'%','輸出負載變動穩壓率CH1_H'),
(45,'%','輸出負載變動穩壓率CH2_L'),
(46,'%','輸出負載變動穩壓率CH2_N'),
(47,'%','輸出負載變動穩壓率CH2_H'),
(48,'mA','短路保護測試_L'),
(49,'mA','短路保護測試_N'),
(50,'mA','短路保護測試_H'),
(51,'V','過電流保護測試_L'),
(52,'V','過電流保護測試_N'),
(53,'V','過電流保護測試_H'),
(54,'V','遠端控制關機測試'),
(55,'V','輸出電壓整理微調正緣'),
(56,'V','輸出電壓整理微調負緣'),
(57,'V','遞減遠端控制關機測試_L'),
(58,'V','遞減遠端控制關機測試_N'),
(59,'V','遞減遠端控制關機測試_H'),
(60,'V','過電壓保護_L'),
(61,'V','過電壓保護_N'),
(62,'V','過電壓保護_H'),
(63,'V','遞減欠壓鎖定_L'),
(64,'A','遞增過電流保護_L'),
(65,'A','遞增過電流保護_N'),
(66,'A','遞增過電流保護_H'),
(67,'V','遞增遠端控制開機測試_L'),
(68,'V','遞增遠端控制開機測試_N'),
(69,'V','遞增遠端控制開機測試_H'),
(70,'%','交越負載CH1_L'),
(71,'%','交越負載CH1_N'),
(72,'%','交越負載CH1_H'),
(73,'%','交越負載CH2_L'),
(74,'%','交越負載CH2_N'),
(75,'%','交越負載CH2_H')


#ctrl - shift - L  同時編輯多行

'空載輸入電流_L'
'空載輸入電流_N'
'空載輸入電流_H'
'空載輸入功率_L'
'空載輸入功率_N'
'空載輸入功率_H'
'空載輸出電壓CH1_L'
'空載輸出電壓CH1_N'
'空載輸出電壓CH1_H'
'空載輸出電壓CH2_L'
'空載輸出電壓CH2_N'
'空載輸出電壓CH2_H'
'空載漣波雜訊CH1_L'
'空載漣波雜訊CH1_N'
'空載漣波雜訊CH1_H'
'空載漣波雜訊CH2_L'
'空載漣波雜訊CH2_N'
'空載漣波雜訊CH2_H'
'滿載輸入電流_L'
'滿載輸入電流_N'
'滿載輸入電流_H'
'滿載輸入功率_L'
'滿載輸入功率_N'
'滿載輸入功率_H'
'滿載輸出電壓CH1_L'
'滿載輸出電壓CH1_N'
'滿載輸出電壓CH1_H'
'滿載輸出電壓CH2_L'
'滿載輸出電壓CH2_N'
'滿載輸出電壓CH2_H'
'滿載漣波雜訊CH1_L'
'滿載漣波雜訊CH1_N'
'滿載漣波雜訊CH1_H'
'滿載漣波雜訊CH2_L'
'滿載漣波雜訊CH2_N'
'滿載漣波雜訊CH2_H'
'滿載效率_L'
'滿載效率_N'
'滿載效率_H'
'輸入電壓變動穩壓率CH1'
'輸入電壓變動穩壓率CH2'
'輸出負載變動穩壓率CH1_L'
'輸出負載變動穩壓率CH1_N'
'輸出負載變動穩壓率CH1_H'
'輸出負載變動穩壓率CH2_L'
'輸出負載變動穩壓率CH2_N'
'輸出負載變動穩壓率CH2_H'
'短路保護測試_L'
'短路保護測試_N'
'短路保護測試_H'
'過電流保護測試_L'
'過電流保護測試_N'
'過電流保護測試_H'
'遠端控制關機測試'
'輸出電壓整理微調正緣'
'輸出電壓整理微調負緣'
'遞減遠端控制關機測試_L'
'遞減遠端控制關機測試_N'
'遞減遠端控制關機測試_H'
'過電壓保護_L'
'過電壓保護_N'
'過電壓保護_H'
'遞減欠壓鎖定_L'
'遞增過電流保護_L'
'遞增過電流保護_N'
'遞增過電流保護_H'
'遞增遠端控制開機測試_L'
'遞增遠端控制開機測試_N'
'遞增遠端控制開機測試_H'
'交越負載CH1_L'
'交越負載CH1_N'
'交越負載CH1_H'
'交越負載CH2_L'
'交越負載CH2_N'
'交越負載CH2_H'

(1,'ma','空載輸入電流_L'),
(2,'ma','空載輸入電流_N'),
(3,'ma','空載輸入電流_H'),
(4,'W','空載輸入功率_L'),
(5,'W','空載輸入功率_N'),
(6,'W','空載輸入功率_H'),
(7,'V','空載輸出電壓CH1_L'),
(8,'V','空載輸出電壓CH1_N'),
(9,'V','空載輸出電壓CH1_H'),
(10,'V','空載輸出電壓CH2_L'),
(11,'V','空載輸出電壓CH2_N'),
(12,'V','空載輸出電壓CH2_H'),
(13,'p-pmV','空載漣波雜訊CH1_L'),
(14,'p-pmV','空載漣波雜訊CH1_N'),
(15,'p-pmV','空載漣波雜訊CH1_H'),
(16,'p-pmV','空載漣波雜訊CH2_L'),
(17,'p-pmV','空載漣波雜訊CH2_N'),
(18,'p-pmV','空載漣波雜訊CH2_H'),
(19,'mA','滿載輸入電流_L'),
(20,'mA','滿載輸入電流_N'),
(21,'mA','滿載輸入電流_H'),
(22,'W','滿載輸入功率_L'),
(23,'W','滿載輸入功率_N'),
(24,'W','滿載輸入功率_H'),
(25,'V','滿載輸出電壓CH1_L'),
(26,'V','滿載輸出電壓CH1_N'),
(27,'V','滿載輸出電壓CH1_H'),
(28,'V','滿載輸出電壓CH2_L'),
(29,'V','滿載輸出電壓CH2_N'),
(30,'V','滿載輸出電壓CH2_H'),
(31,'p-pmV','滿載漣波雜訊CH1_L'),
(32,'p-pmV','滿載漣波雜訊CH1_N'),
(33,'p-pmV','滿載漣波雜訊CH1_H'),
(34,'p-pmV','滿載漣波雜訊CH2_L'),
(35,'p-pmV','滿載漣波雜訊CH2_N'),
(36,'p-pmV','滿載漣波雜訊CH2_H'),
(37,'%','滿載效率_L'),
(38,'%','滿載效率_N'),
(39,'%','滿載效率_H'),
(40,'%','輸入電壓變動穩壓率CH1'),
(41,'%','輸入電壓變動穩壓率CH2'),
(42,'%','輸出負載變動穩壓率CH1_L'),
(43,'%','輸出負載變動穩壓率CH1_N'),
(44,'%','輸出負載變動穩壓率CH1_H'),
(45,'%','輸出負載變動穩壓率CH2_L'),
(46,'%','輸出負載變動穩壓率CH2_N'),
(47,'%','輸出負載變動穩壓率CH2_H'),
(48,'mA','短路保護測試_L'),
(49,'mA','短路保護測試_N'),
(50,'mA','短路保護測試_H'),
(51,'V','過電流保護測試_L'),
(52,'V','過電流保護測試_N'),
(53,'V','過電流保護測試_H'),
(54,'V','遠端控制關機測試'),
(55,'V','輸出電壓整理微調正緣'),
(56,'V','輸出電壓整理微調負緣'),
(57,'V','遞減遠端控制關機測試_L'),
(58,'V','遞減遠端控制關機測試_N'),
(59,'V','遞減遠端控制關機測試_H'),
(60,'V','過電壓保護_L'),
(61,'V','過電壓保護_N'),
(62,'V','過電壓保護_H'),
(63,'V','遞減欠壓鎖定_L'),
(64,'A','遞增過電流保護_L'),
(65,'A','遞增過電流保護_N'),
(66,'A','遞增過電流保護_H'),
(67,'V','遞增遠端控制開機測試_L'),
(68,'V','遞增遠端控制開機測試_N'),
(69,'V','遞增遠端控制開機測試_H'),
(70,'%','交越負載CH1_L'),
(71,'%','交越負載CH1_N'),
(72,'%','交越負載CH1_H'),
(73,'%','交越負載CH2_L'),
(74,'%','交越負載CH2_N'),
(75,'%','交越負載CH2_H'),

