# coding: UTF-8
import pandas as pd
import time
import sys
import os
import math
import numpy as np
from pymeasure.instruments.keithley import Keithley2400
from datetime import datetime
import concurrent.futures

### thermistor settings ###
pulse_interval = 0.4
pulse_current = 0.000001 # 1uA

#thermistor 104JT-025
B_constant = 4390
R25 = 100000
T25 = 298.15

keithley1 = Keithley2400("GPIB::22") # thermistor
keithley2 = Keithley2400("GPIB::24") # current_apply

voltage_interval = 0.2

def initial_settings():
    ### keithley settings ###
    keithley1.reset()
    keithley1.disable_buffer()
    keithley1.use_front_terminals()
    keithley1.apply_current()
    keithley1.source_voltage_range = 0.1
    keithley1.source_current = pulse_current
    keithley1.enable_source()
    keithley1.measure_voltage()
    keithley1.compliance_voltage = 1.8

    keithley2.reset()    
    keithley2.disable_buffer()
    keithley2.use_front_terminals()
    keithley2.apply_current()
    keithley2.source_voltage_range = 0.3   
    keithley2.source_current = 0
    keithley2.enable_source()
    keithley2.measure_voltage()
    keithley2.compliance_voltage = 1.3 # limit voltage near the potential window of the water

# print the estimated time (h/m/s) for measurement before starting
def time_calculation():
    # in the unit of second
    total_required_time = before_measurement_h * 3600 + measurement_interval_h * 3600 * (measurements - 1) + measurements   
    for j in range(measurements) :
        total_required_time = total_required_time + (period_number[j] + 1) * half_period_time[j] * 2  
    
    # convert into (h/m/s)
    required_time_h = int(total_required_time/3600)
    required_time_m = int((total_required_time-required_time_h*3600)/60)
    required_time_s = int(total_required_time -(required_time_h*3600+required_time_m*60))
    print('total required_time: ', required_time_h,' h ',required_time_m,' m ',required_time_s,' s ')

def ec_confirmation():
    print('sample name:', sample_name)
    print('\n')
    print('before measurement:', before_measurement_h,'h')
    for i in range(measurements):
        print('experiment', i+1)
        print('    ',I[i],'mA,',half_period_time[i],'s,',period_number[i],'+1 periods')
    print('measurement interval:',measurement_interval_h,'h')

def connection_test():
    keithley1.reset()
    keithley1.disable_buffer()
    keithley1.use_front_terminals()
    keithley1.apply_current()
    keithley1.source_current = 0
    keithley1.enable_source()

    keithley2.reset()
    keithley2.disable_buffer()
    keithley2.use_front_terminals()
    keithley2.apply_current()
    keithley2.source_current = 0
    keithley2.enable_source()

    time.sleep(0.5)
    keithley1.shutdown()
    keithley2.shutdown()    

    print('connection OK')
    
# apply current to a thermistor and get a voltage value
def thermistor():
    target_time = 0
    time_accuracy = 0.001 # time for sleep
    rest_until_target_time = time.time() - base_time - target_time
    for step in range(thermistor_steps):
        while rest_until_target_time < 0: # wait until reaching target time
            time.sleep(time_accuracy) #wait for time_accuracy, then recalculate target time
            rest_until_target_time = time.time() - base_time - target_time
        keithley1.source_current = pulse_current # when reaching the target time, apply current
        keithley1.ramp_to_current(pulse_current,steps=1, pause = 0.000001)
        t2 = time.time()
        V = keithley1.voltage
        keithley1.source_current = 0 # after obtaining data turn off the current immediately

        R = V/pulse_current
        temperature = 1/(1/T25 + math.log(R/R25)/B_constant) - 273.15
        sheet.cell(row = step+2, column = 1).value = t2 - base_time
        sheet.cell(row = step+2, column = 2).value = temperature # calculate a temperature from a voltage value, and record to the excel file

        target_time = target_time + pulse_interval
        rest_until_target_time = time.time() - base_time - target_time #reset the target time
    keithley1.shutdown()
    print('thermistor_done.')

# apply a current (alternating square wave current) to the ECP cell 
def current_apply(current):
    target_time = half_period_time[i]
    voltage_time = voltage_interval
    time_accuracy = 0.001

    rest_until_target_time = time.time() - base_time - target_time
    for step in range((period_number[i]+1) *2):
        keithley2.ramp_to_current(current,steps=3)
        while rest_until_target_time < 0: # wait until the time to reverse the current direction
            rest_voltage = time.time()-base_time - voltage_time
            while rest_voltage<0 and rest_until_target_time<0: # recird voltage vales, priotizing current flipping
                time.sleep(time_accuracy)
                rest_voltage = time.time()-base_time - voltage_time
                rest_until_target_time = time.time() - base_time - target_time
            sheet.cell(row = int(voltage_time/voltage_interval)+1, column = 12).value = time.time()-base_time
            sheet.cell(row = int(voltage_time/voltage_interval)+1, column = 13).value = keithley2.voltage # record to the excel file
            voltage_time = voltage_time + voltage_interval #reset the time to record voltage
            rest_until_target_time = time.time() - base_time - target_time
        current = current*(-1)
        target_time = target_time + half_period_time[i]
        rest_until_target_time = time.time() - base_time - target_time

    # at last, apply extra current for 1 sec to justify the calculation of Tox-Tred(最後に、1秒だけ余分に電流を流す。Tox-Tredの計算の帳尻を合わせるため。)
    keithley2.ramp_to_current(current,steps=3)
    time.sleep(1)
    keithley2.shutdown()
    sheet.cell(row = 11, column = 5).value = voltage_interval # you can save voltage_interval outside of the function
    sheet.cell(row = 6, column = 5).value = current*1000 # record current to the excel file (current was difined in this function)
                                                        
    print('current_apply_done.')

def calculation():
    starting_cell = 2 # row number of the cell corresponding to 0 s
#    starting_cell_calc = starting_cell + half_period_points*2 # row number of the cell corresponding to 0 s
#    print("test, starting_cell:", starting_cell_calc)

    ### average temperature of each points within the cycle###
    sheet.cell(row = 1, column = 7).value = 't (s)'
    for a in range(0,half_period_points*2):
        sheet.cell(row = starting_cell+a, column = 7).value = a*pulse_interval
    sheet.cell(row = 1, column = 8).value = 'Average temperature (C)'
    for b in range(0,half_period_points*2): 
        total = 0
        row_number = 0
        for step in range (0,period_number[i]):
            total = total + (sheet.cell(row = (starting_cell+half_period_points*2) + row_number + b, column = 2).value) # summation 
            #1st cycle is rejected
            row_number = row_number + half_period_points*2 # hop n to n+1, by using row_number
        #print((b,",","total:",total))
        sheet.cell(row = starting_cell + b, column = 8).value = total/period_number[i] #averaged

    ### time of 1 cycle ###
    sheet.cell(row = 1, column = 9).value = 't (s)'
    for a in range(0,half_period_points):
        sheet.cell(row = starting_cell+a, column = 9).value = a*pulse_interval # 例えば0.2秒感覚で測定したなら、0.2, 0.4, 0.6...と各セルに入力。

    ### Tox(t) - Tred(t) ### red側につないだとき、つまりサーミスタのついた電極では酸化→還元の順でサイクルが回っているときを想定。
    sheet.cell(row = 1, column = 10).value = 'Tox(t)-Tred(t) (K)'
#    half_period_points = int(half_period_time[i]/pulse_interval) #次のfor文のために、Tox-Tredの計算に必要な点数を計算。
    for c in range(0,half_period_points):
        sheet.cell(row = starting_cell + c, column = 10).value = sheet.cell(row = starting_cell + c, column = 8).value - sheet.cell(row = starting_cell+ c + half_period_points, column = 8).value 
        # 酸化→還元の順で起こるので、Average emperatureの列から、酸化開始a秒後-還元開始a秒後(a<=half_period_time)をそれぞれ計算。

    print('cauculation_done.')


### conducting functions ###
connection_test()
while True:
    meas_No = input("How many times do you plan to measure?:") #measurement number
    try:
        if meas_No == "q":
            print("Cancelled")
            break
        meas_No = int(meas_No)
        sample_names = []
        currents = []
        cycles = []
        flip_times = []
        i = 0
        while i < meas_No:
            print("Cycle No" + str(i + 1) + ".\n") 
            sample_name = input("Put your sample name:")
            current = input("Current [mA]:")
            cycle = input("How many cycles?:")
            flip_time = input("How long [s] will the cycle?:")
            if sample_name == "q" or current == "q" or cycle == "q" or flip_time == "q":
                print("cancelled.")
                break
            sample_names.append(sample_name)
            currents.append(float(current))
            cycles.append(int(cycle))
            flip_times.append(float(flip_time) / 2)
            while True:
                confirm = input("Are you sure? [y/n]:")
                if confirm == "y":
                    print("OK.")
                    i += 1
                    break
                elif confirm == "n":
                    sample_names.pop()
                    currents.pop()
                    cycles.pop()
                    flip_times.pop()
                    print("Retry")
                    break
                else:
                    print("?")
                
        wait_time = float(input("How much time [h] should passe before measurements?:")) * 3600
        int_time = float(input("Time between measurements [h]:")) * 3600 # interval time
        current_cycle = 0
        while current_cycle < len(sample_names):
            print("/nNow cycle" + str(current_cycle + 1) + "is being conducted.\n")
            data = pd.DataFrame(columns=["time(V)[V]","V[V]","time(T)[C]","Temp[c]","cycle_time[s]","ave_Temp[C]"])
            
    
            
    except:
        if 
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(exc_type, exc_value, "[in line",exc_tb.tb_lineno,"]")

    ender = input("q to end.:")
    if ender == "q":
        print("Bye.")
        break
        
if __name__ == "__main__":
    sample_name = 'ferri_ferrocyanide_0.4M'
    I = np.array([0.1,0.2,0.3,0.4]) # 各測定の電流値をmA単位で入力。numpyのリストにしておくと1000で割って容易にA単位に直せる。
    period_number = [5000,5000,5000,5000,1,1] #始めが5000だったwakayama
    half_period_time = [4,4,4,4,4,4]
    before_measurement_h = 8 # waiting time for thermal equilibrium, at least 3 (unit:hours).
    measurement_interval_h = 0 # 測定ごとの間隔(hours)。ここを「2」にすると、測定ごとに2時間間隔を空ける。    

    measurements = int(input('number of measrurements: ')) # 繰り返し測定の回数を入力。

    # ユーザ名メモ。Peltier, kimilb2019_8, fmato
    # C:\Users\kimilab2019_8\Desktop\filename.xlsx

    time_calculation() # 測定時間をhmsで計算し、表示。
    ec_confirmation()
    confirmation=int(input('conduct measuremant : 1')) # 測定条件に問題なければ1、問題あれば1以外の文字を入力。

    np.array(I, dtype=float)# floatリストに変換
    I = I/1000 # 電流値をA単位へ変換
    np.array(period_number, dtype=int) # integerリストに変換
    np.array(half_period_time, dtype=float)# floatリストに変換

    if confirmation == 1: #測定条件に相違なければ繰り返し測定を実行。
        connection_test()
        print(datetime.now())
        time.sleep(before_measurement_h*3600)
        day = datetime.now().strftime("%Y_%m_%d")
        path_peltier = "C:/Users/Hiroshi/Documents/data/peltier/"
        path = path_peltier + day
        if os.path.exists(path) == False:
            os.chdir(path_peltier)
            os.mkdir(day)
        else:
            pass
        for i in range(measurements):
            print(datetime.now())
            step = 0
            book = openpyxl.Workbook() # エクセルファイルを作成
            sheet = book.worksheets[0] 
            initial_settings() #ソースメータの電源を入れる
            half_period_points = int(half_period_time[i]/pulse_interval) 
            required_time = (period_number[i]+1) * half_period_time[i] *2 #測定時間（秒）を計算。 
            thermistor_steps = int(required_time/pulse_interval)+int(1/pulse_interval) 

            sheet.cell(row = 7, column = 5).value = period_number[i]+1
            sheet.cell(row = 8, column = 5).value = half_period_time[i]

            base_time = time.time()
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=2) # サーミスタと電流引加をマルチスレッドで実行。マルチプロセスだとなぜか上手く動かない。
            executor.submit(thermistor)
            executor.submit(current_apply(I[i]))
            executor.shutdown() #両方の操作が終わったら次の操作へ。
            date = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
            book.save("{0}/{2}_{3}mA_{4}s_{5}+1_{1}.xlsx".format(path, date, sample_name, I[i]*1000, half_period_time[i], period_number[i]))
            calculation()
            #chart_creation()
            book.save("{0}/{2}_{3}mA_{4}s_{5}+1_{1}.xlsx".format(path, date, sample_name, I[i]*1000, half_period_time[i], period_number[i]))
            if i == measurements-1:
                break
            time.sleep(measurement_interval_h * 3600)
        print('finished.')

    else: 
    #設定ミスなどがあれば測定前に測定をキャンセルする。
    # jupyter lab側から測定を中止するのはなぜか時間がかかったり再起動が必要だったりしてとにかく面倒。
        print('cancelled')
