from __future__ import annotations
import subprocess


class Simulation:
    """Connects to a given simulation.
        
    Args:
        path (str): String with the raw path to the Aspen PLUS file.
    """ 
    def __init__(self, path: str) -> None:     
        
        import win32com.client as win32  
        
        self.path = path
        self.case = win32.Dispatch('Apwn.Document')
        self.case.InitFromArchive2(path)
        
    def set_visible(self, visibility:int = 0) -> None:
        """Sets the visibility of the flowsheet.
        
        Args:
            visibility (int, optional): If 1, it shows the flowsheet. If 0, it keeps it invisible.. Defaults to 0.
        """
        self.case.Visible = visibility
        
    def set_popups(self, popups:bool = True) -> None:
        """Allows or supresses Aspen popups. For example, when reinitializing, it asks if
        you are sure.

        Args:
            popups (bool, optional): State if you want or not popups. Defaults to True.
        """
        self.case.SuppressDialogs = popups
    
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
        I do not appreciate that. If you do not want it to do this, set soft = True.
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
    def BLOCK(self) -> 'TreeBlock':
        """Shortcut to access the blocks in the Aspen variable tree.

        Returns:
            TreeBlock: Element of the variable tree correspoonding to Data\Blocks
        """                
        return self.case.Tree.Elements("Data").Elements("Blocks")
    @property
    def STREAM(self) -> 'TreeStream':
        """Shortcut to access the streams in the Aspen variable tree.

        Returns:
            TreeStream: Element of the variable tree corresponding to Data\Streams
        """        
        return self.case.Tree.Elements("Data").Elements("Streams")
    
    def get_stream(self, name: str) -> ProcessStream:
        """Get a stream object

        Args:
            name (str): Name of the stream

        Returns:
            ProcessStream: Stream object
        """        
        stream = self.STREAM.FindNode(name)
        if stream:
            return ProcessStream(stream)
        else:
            print(f"There is no stream with name {name} in the simulation")
    
    def get_block(self, name:str) -> ProcessBlock:
        """Get a block object

        Args:
            name (str): Name of the block

        Returns:
            ProcessBlock: Block object
        """     
        block = self.BLOCK.FindNode(name)
        if block:
            return ProcessBlock(block)
        else: 
            print(f"There is no block with name {name} in the simulation")   

class ProcessStream:
    """Creates a process stream from a node in a simulation

    Args:
        node ("AspenObject"): COM object for a node
    """
    def __init__(self, node: "AspenObject") -> None:
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
                      "MASSENTHALPY": r"HMX_MASS\MIXED",
                      "MOLEENTHALPY": r"HMX\MIXED",
                      }
        for property in prop_list:
            if type(property) == tuple:
                component = "\\" + property[1]
                property_key = property[0]
            else:
                component = ""
                property_key = property
            if property_key in match_dict:
                try:
                    properties[property] = self.stream.FindNode(r"Output\{0}{1}".format(match_dict[property_key], component)).Value
                except AttributeError:
                    print(f"ERROR: Something went wrong with how property {property_key} is looked. Check manually")
            else:
                print(f"WARNING: Property {property} not found. May not be implemented, or doesn't exist")
        return properties
    
    def set_properties(self, prop_dict:dict) -> None:
        """Set the stream properties using a dictionary

        Args:
            prop_dict (dict): Dict of properties {key,value}. The valid elements of the keys are:
            "TEMP": Temperature
            "PRES": Pressure 
            "FLOW": Total flow. Depends on FLOWBASIS
            "FLOWBASIS": Basis of the total flow. Allowed values are "MASS", "MOLE", "STDVOL", or "VOLUME"
            ("COMFLOW", "chemical"): Component flow. Depends on COMPBASIS
            "COMPBASIS": Basis of the composition window. It can be "MASS-FLOW", "MOLE-FLOW", "STDVOL-FLOW", "MASS-FRAC", "MOLE-FRAC", "STDVOL-FRAC", "MASS-CONC" or"MOLE-CONC" 
            "VAPFRAC": Vapor fraction
            "FLASHTYPE": Type of the data that the streams requires. Options are "TP", "TV", or "PV", where P is pressure, T is temperature, and V is vapor fraction
        """
        match_dict = {"TEMP": r"TEMP\MIXED",
                      "PRES": r"PRES\MIXED",
                      "FLOW": r"TOTFLOW\MIXED",
                      "FLOWBASIS": r"FLOWBASE\MIXED", 
                      "COMPFLOW": r"FLOW\MIXED",
                      "COMPBASIS": r"BASIS\MIXED", 
                      "VAPFRAC": r"VFRAC\MIXED",
                      "FLASHTYPE": r"MIXED_SPEC\MIXED"
                    }
        for property in prop_dict:
            if type(property) == tuple:
                component = "\\" + property[1]
                property_key = property[0]
            else:
                component = ""
                property_key = property
            if property_key in match_dict:
                try:
                    self.stream.FindNode(r"Input\{0}{1}".format(match_dict[property_key], component)).Value = prop_dict[property]
                except AttributeError:
                    print(f"ERROR: Something went wrong with how property {property_key} is set. Check manually.")
            else:
                print(f"WARNING: Property {property} not found. May not be implemented, or doesn't exist")
        
class ProcessBlock:
    """Creates a process block from a node in a simulation

    Args:
        node (AspenObject): COM object for a block
    """            
    def __init__(self, node:"AspenObject") -> None:
        self.block = node
            
        
           
        