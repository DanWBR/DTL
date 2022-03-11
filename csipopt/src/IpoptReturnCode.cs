// Copyright (C) 2010-2011 Anders Gustafsson and others. All Rights Reserved.
// This code is published under the Eclipse Public License.
//
// Author:  Anders Gustafsson, Cureos AB 2011-12-13

namespace Cureos.Numerics
{
    /// <summary>
    /// Return codes for the Solve call for an IPOPT problem 
    /// </summary>
    public enum IpoptReturnCode
    {
#pragma warning disable 1591
        Solve_Succeeded = 0,
        Solved_To_Acceptable_Level = 1,
        Infeasible_Problem_Detected = 2,
        Search_Direction_Becomes_Too_Small = 3,
        Diverging_Iterates = 4,
        User_Requested_Stop = 5,
        Feasible_Point_Found = 6,

        Maximum_Iterations_Exceeded = -1,
        Restoration_Failed = -2,
        Error_In_Step_Computation = -3,
        Maximum_CpuTime_Exceeded = -4,
        Not_Enough_Degrees_Of_Freedom = -10,
        Invalid_Problem_Definition = -11,
        Invalid_Option = -12,
        Invalid_Number_Detected = -13,

        Unrecoverable_Exception = -100,
        NonIpopt_Exception_Thrown = -101,
        Insufficient_Memory = -102,
        Internal_Error = -199,

        Problem_Not_Initialized = -900
#pragma warning restore 1591
    }
}