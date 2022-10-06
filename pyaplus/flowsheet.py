import win32com.client as win32
import subprocess

class Simulation:
    """Connects to a given simulation.
        
    Args:
        path (str): String with the raw path to the Aspen PLUS file. If "Active", chooses the open HYSYS flowsheet.
    """ 
    def __init__(self, path: str) -> None:       
        self.path = path
        self.case = win32.Dispatch('Apwn.Document')
        self.case.InitFromArchive2(path)
        
    def set_visible(self, visibility:int = 0) -> None:
        """Sets the visibility of the flowsheet.
        
        Args:
            visibility (int, optional): If 1, it shows the flowsheet. If 0, it keeps it invisible.. Defaults to 0.
        """        
        self.case.Visible = visibility
    
    def run(self) -> None:
        """Runs the simulation.
        """        
        self.case.Run()
        
    def reinit(self) -> None:
        """Reinitiates the simulation
        """        
        self.case.Reinit()

    def close(self, soft: bool = False) -> None:
        """Closes the instance and the Aspen connection. If you do not close it,
        the task will remain and you will have to stop it from the task manage.
        
        WARNING: It will close ALL Aspen Plus instances. Cause it was not being obedient. And
        I do not appreciate that. If you do not want it to do this, set soft = True
        """        
        self.case.Close(self.path)
        self.case.Quit()
        print("Aspen Case was closed")
        if not soft:
            cmd = "WMIC PROCESS where name='AspenPlus.exe' get Caption,Commandline,Processid"
            proc = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE)
            list_aspen_processes = [i for i in proc.stdout]
            if len(list_aspen_processes) > 2:
                cmd = "wmic process where name='AspenPlus.exe' call terminate"
                subprocess.Popen(cmd, shell=False, stdout = subprocess.DEVNULL)
                print("FORCEFULLY!")
        del self
        
    @property
    def BLOCK(self):        
        return self.case.Tree.Elements("Data").Elements("Blocks")
    @property
    def STREAM(self):
        return self.case.Tree.Elements("Data").Elements("Streams")
    
    def get_stream(self, name: str) -> any:
        """Get a stream object

        Args:
            name (str): Name of the stream

        Returns:
            StreamObject: Stream object
        """        
        stream = self.STREAM.FindNode(name)
        if stream:
            return ProcessStream(stream)
        else:
            print(f"There is no stream with name {name} in the simulation")
    
    def get_block(self, name:str) -> any:
        """Get a block object

        Args:
            name (str): _description_

        Returns:
            BlockObject: _description_
        """     
        block = self.BLOCK.FindNode(name)
        if block:
            return ProcessBlock(block)
        else: 
            print(f"There is no block with name {name} in the simulation")   

class ProcessStream:
    """Creates a process stream from a node in a simulation

    Args:
        node (any): COM object for a node
    """
    def __init__(self, node: any) -> None:
        self.stream = node 
        
    def get_properties(self, prop_list:list) -> dict:
        """Gets the stream properties in a dictionary

        Args:
            prop_list (list): List of properties. The valid elements of the list are:\n
            "TEMP": Temperature\n
            "PRES": Pressure\n
            "MOLEFLOW": Molar flow\n
            ("COMPMOLEFLOW", "chemical"): Component molar flow of a chemical\n
            "MASSFLOW": Mass flow\n
            ("COMPMASSFLOW", "chemical"): Component mass flow of a chemical\n
            ("COMPMASSFRAC", "chemical"): Component mass fraction of a chemical\n
            ("COMPMOLEFRAC", "chemical"): Component mole fraction of a chemical\n
            "VOLUMETRICFLOW": Volumetric flow\n
            "MASSENTHALPY": Enthalpy per unit of mass\n
            "MOLEENTHALPY": Enthalpy per mole unit
        
        Returns:
            dict: Dictionary with the property and the value. The units have to be checked in the Aspen Plus simulation flowsheet.
        """        
        properties = {}
        match_dict = {"TEMP": r"TEMP_OUT\MIXED",
                      "PRES": r"PRES_OUT\MIXED",
                      "MOLEFLOW": r"MOLEFLMX\MIXED",
                      "COMPMOLEFLOW": r"MOLEFLOW\MIXED",
                      "MASSFLOW": r"MASSFLMX\MIXED",
                      "COMPMASSFLOW": r"MASSFLOW\MIXED",
                      "COMPMOLEFRAC": r"MOLEFRAC\MIXED",
                      "COMPMASSFRAC": r"MASSFRAC\MIXED",
                      "VOLUMETRICFLOW": r"VOLFLMX\MIXED",
                      "MASSENTHALPY": r"HMXMASS\MIXED",
                      "MOLEENTHALPY": r"HMX\MIXED",
                      }
        for property in prop_list:
            if type(property) == tuple:
                component = "\\" + property[1]
                property = property[0]
            else:
                component = ""
            if property in match_dict:
                properties[property + component] = self.stream.FindNode(r"Output\{0}{1}".format(match_dict[property], component)).Value
            else:
                print(f"Property {property} not found. May not be implemented, or doesn't exist")
        return properties
    
class ProcessBlock:
    """Creates a process block from a node in a simulation

    Args:
        node (any): COM object for a block
    """            
    def __init__(self, node:any) -> None:
        self.block = node
            
        
           
        