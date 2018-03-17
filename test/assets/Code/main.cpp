/**
 * name - main.cpp
 *
 * \file
 */

#include "constants.h"
#include "CommonUtility.h"
#include "utils.h"

/**
 * name: main
 *
 * \REQUIREMENT_LINK Req 3A
 */
int main(int argc, char* argv[])
{
    CommonUtility* itsCommonUtility = new CommonUtility();

    int value = CommonUtility::iGetConstant();

    // increment value
    value -= utilityA();
    // increment value again
    utilityB(&value);

    // print constant
    itsCommonUtility->iPrintConstant(value);

    return 0;
}
