from enum import Enum


class Component(Enum):
    CAPACITANCE = 0
    RESISTANCE = 1
    OTHER = 2
    NONE = 3


class Common(object):
    @staticmethod
    def isComponent(value, voltage, footprint):
        component = Component.OTHER
        value = str(value)
        if len(value) != 0 and value.rstrip()[-1].upper() == 'F':
            if voltage.rstrip()[-1].upper() == 'V':
                component = Component.CAPACITANCE  # is capacitance
        elif value.upper().rfind('K') != -1 or value.upper().rfind('R') != -1 or value.upper().rfind('M') != -1:
            if len(footprint.rstrip()) == 5:
                if footprint.lstrip()[0] == 'R':
                    component = Component.RESISTANCE  # is resistance
        return component

