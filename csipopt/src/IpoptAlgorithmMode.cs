// Copyright (C) 2010-2011 Anders Gustafsson and others. All Rights Reserved.
// This code is published under the Eclipse Public License.
//
// Author:  Anders Gustafsson, Cureos AB 2011-12-13

namespace Cureos.Numerics
{
    /// <summary>
    /// Enumeration to indicate the mode in which the algorithm is
    /// </summary>
    public enum IpoptAlgorithmMode
    {
#pragma warning disable 1591
        RegularMode = 0,
        RestorationPhaseMode = 1
#pragma warning restore 1591
    }
}