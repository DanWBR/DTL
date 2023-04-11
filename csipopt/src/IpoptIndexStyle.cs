// Copyright (C) 2010-2011 Anders Gustafsson and others. All Rights Reserved.
// This code is published under the Eclipse Public License.
//
// Author:  Anders Gustafsson, Cureos AB 2011-12-14

namespace Cureos.Numerics
{
    /// <summary>
    /// Enumeration of the available index styles for the Jacobian and Hessian index arrays.
    /// </summary>
    public enum IpoptIndexStyle
    {
#pragma warning disable 1591
        C = 0,
        Fortran = 1
#pragma warning restore 1591
    }
}