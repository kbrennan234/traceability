/**
 * name - CommonUtility.cpp
 *
 * \file
 */

#include "CommonUtility.h"
#include "constants.h"
#include <assert.h>
#include <stdio.h>

CommonUtility::CommonUtility()
{
}

CommonUtility::~CommonUtility()
{
}

int CommonUtility::iGetConstant()
{
    /// \REQUIREMENT_LINK Req 5A
    static_assert(0 == CONSTANT_A, "Expected constant to be 0");

    /// \REQUIREMENT_LINK Req 5A
    return CONSTANT_A;
}

void CommonUtility::iPrintConstant(int value)
{
    /**
     * \REQUIREMENT_LINK Req 6B
     * \REQUIREMENT_LINK Req 7B
     */
    printf("\n\nThis is your value: %d\n\n", value);
}
