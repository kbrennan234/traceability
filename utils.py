from collections import namedtuple
import enum

''' Requirement details '''
tRequirementValue = namedtuple('tRequirementValue', ['reqText', 'reqLinks'])

''' Requirement link details '''
tRequirementLink = namedtuple('tRequirementLink', ['linkType', 'linkName', 'linkFile', 'linkFileLineNum'])
# override link equality operator
tRequirementLink.__eq__ = lambda x, y: \
    (x.linkType == y.linkType) and \
    (x.linkName == y.linkName) and \
    (x.linkFile == y.linkFile) and \
    (x.linkFileLineNum == y.linkFileLineNum)

class tLinkType(enum.Enum):
    ''' Requirement link types'''
    LINK_TYPE__SRC = 1
    LINK_TYPE__TEST = 2