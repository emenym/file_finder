from easygui import EgStore
#-----------------------------------------------------------------------
# define a class named Settings as a subclass of EgStore
#-----------------------------------------------------------------------
class Settings(EgStore):

    def __init__(self, filename):  # filename is required
        #-------------------------------------------------
        # Specify default/initial values for variables that
        # this particular application wants to remember.
        #-------------------------------------------------
        # self.userId = ""
        self.targetServer = ""
        self.firstName = ""
        self.lastName = ""

        #-------------------------------------------------
        # For subclasses of EgStore, these must be
        # the last two statements in  __init__
        #-------------------------------------------------
        self.filename = filename  # this is required
        self.restore()            # restore values from the storage file if possible