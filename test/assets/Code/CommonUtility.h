/**
 * name - CommonUtility.h
 *
 * \file
 */

#ifndef CommonUtility_H
#define CommonUtiity_H

/**
 * name - CommonUtility
 * description - Utility class for common functionality
 *
 * \REQUIREMENT_LINK Req 1A
 * \REQUIREMENT_LINK Req 2A
 * \REQUIREMENT_LINK Req 3A
 */
class CommonUtility
{
public:

    CommonUtility();

    ~CommonUtility();

    static int iGetConstant();

    void iPrintConstant(int value);

protected:

    /// \REQUIREMENT_LINK Req 7B
    static const int UTILITY_CONSTANT = 1234;
};

#endif
