from time import sleep
import wmi
import sys

'''
插上电源AC后，放电率会等于0,打破
while True回圈，打印最终结果
'''

c = wmi.WMI()
t = wmi.WMI(moniker="//./root/wmi")

energyConsumed = 0

samples = 0
isBegin = False


def getBatInfo():
    full = 0
    percent = 0

    batts = t.ExecQuery('Select * from BatteryFullChargedCapacity')
    for i, b in enumerate(batts):
        full = b.FullChargedCapacity

    batts = c.ExecQuery('Select * from Win32_Battery')
    for i, b in enumerate(batts):
        percent = b.EstimatedChargeRemaining

    batts = t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
    for i, b in enumerate(batts):
        return (b.DischargeRate, b.Voltage, b.RemainingCapacity, full, percent)


beginCapability = 0
oriFullCapability = 0
energyPercentage = None
dischargeRate = 0  # 放电率
voltage = 0  # 电压
systemCapacity = 0
calculatedCapacity = 0
while True:
    sleep(1)
    res = getBatInfo()
    if isBegin:
        if res[0] == 0:
            break
        samples = samples+1
        energyConsumed = energyConsumed+res[0]/3600
        energyPercentage = res[4]
        dischargeRate = res[0]
        voltage = res[1]/1000
        systemCapacity = res[2]
        calculatedCapacity = res[3]
        sys.stdout.write("\r Sample: #%d  (%d%%)  \033[36mEC:\033[0m %.2f mWh  \033[32mDR:\033[0m %d mW  \033[33mV:\033[0m %.3f v \
         \033[34mSC:\033[0m %d/%d mWh \033[35mCC:\033[0m %.2f/%d mWh     " % (samples, energyPercentage, energyConsumed, dischargeRate, voltage, res[2], res[3], beginCapability-energyConsumed, res[3]))
        sys.stdout.flush()
    elif res[0] > 0:
        isBegin = True
        beginCapability = res[2]
        oriFullCapability = res[3]
print()
print("A/C Adapter detected, Wait 5s for capacity update...")
sleep(5)
res = getBatInfo()
print("Full Charge Capacity: %d -> %d mWh" % (oriFullCapability, res[3]))
