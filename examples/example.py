from pyaplus.flowsheet import Simulation

FS = Simulation(path = r"AspenFlowsheet.bkp")

# Assuming you have a stream called S1
S1 = FS.get_stream("S1")
S1.properties = S1.get_properties(["TEMP", ("COMPMOLEFLOW", "CO2"), "MOLEFLOW", "PRES"])
print(S1.properties)

# Closing. If the message FORCEFULLY appears, it means it closes all Aspen Plus instances, cause the API likes to not behave
FS.close()