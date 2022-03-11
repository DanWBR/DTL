// Copyright (C) 2010-2011 Anders Gustafsson and others. All Rights Reserved.
// This code is published under the Eclipse Public License.
//
// Author:  Anders Gustafsson, Cureos AB 2011-12-12

using System;
using System.Runtime.InteropServices;

namespace Cureos.Numerics
{
	#region CALLBACK FUNCTION DELEGATES

	/// <summary>
	/// Delegate defining the callback function for evaluating the value of
	/// the objective function.  Return value should be set to false if
	/// there was a problem doing the evaluation.
	/// </summary>
	/// <param name="n">Number of problem variables</param>
	/// <param name="x">Problem variables</param>
	/// <param name="new_x">true if problem variables are new for this call, false otherwise</param>
	/// <param name="obj_value">Evaluated objective function value</param>
	/// <param name="p_user_data">Optional pointer to user defined data</param>
	/// <returns>true if evaluation succeeded, false otherwise</returns>
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	public delegate IpoptBoolType Eval_F_CB(int n, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] double[] x,
	IpoptBoolType new_x, out double obj_value, IntPtr p_user_data);

	/// <summary>
	/// Delegate defining the callback function for evaluating the gradient of
	/// the objective function.  Return value should be set to false if
	/// there was a problem doing the evaluation.
	/// </summary>
	/// <param name="n">Number of problem variables</param>
	/// <param name="x">Problem variables</param>
	/// <param name="new_x">true if problem variables are new for this call, false otherwise</param>
	/// <param name="grad_f">Objective function gradient</param>
	/// <param name="p_user_data">Optional pointer to user defined data</param>
	/// <returns>true if evaluation succeeded, false otherwise</returns>
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	public delegate IpoptBoolType Eval_Grad_F_CB(int n, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] double[] x,
	IpoptBoolType new_x, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] double[] grad_f, IntPtr p_user_data);

	/// <summary>
	/// Delegate defining the callback function for evaluating the value of
	/// the constraint functions.  Return value should be set to false if
	/// there was a problem doing the evaluation.
	/// </summary>
	/// <param name="n">Number of problem variables</param>
	/// <param name="x">Problem variables</param>
	/// <param name="new_x">true if problem variables are new for this call, false otherwise</param>
	/// <param name="m">Number of constraint functions</param>
	/// <param name="g">Calculated values of the constraint functions</param>
	/// <param name="p_user_data">Optional pointer to user defined data</param>
	/// <returns>true if evaluation succeeded, false otherwise</returns>
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	public delegate IpoptBoolType Eval_G_CB(int n, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] double[] x,
	IpoptBoolType new_x, int m, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 3)] double[] g, IntPtr p_user_data);

	/// <summary>
	/// Delegate defining the callback function for evaluating the Jacobian of
	/// the constraint functions.  Return value should be set to false if
	/// there was a problem doing the evaluation.
	/// </summary>
	/// <param name="n">Number of problem variables</param>
	/// <param name="x">Problem variables</param>
	/// <param name="new_x">true if problem variables are new for this call, false otherwise</param>
	/// <param name="m">Number of constraint functions</param>
	/// <param name="nele_jac">Number of non-zero elements in the Jacobian</param>
	/// <param name="iRow">Row indices of the non-zero Jacobian elements, defined here if values is null</param>
	/// <param name="jCol">Column indices of the non-zero Jacobian elements, defined here if values is null</param>
	/// <param name="values">Values of the non-zero Jacobian elements</param>
	/// <param name="p_user_data">Optional pointer to user defined data</param>
	/// <returns>true if evaluation succeeded, false otherwise</returns>
#if MONO
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	public delegate IpoptBoolType Eval_Jac_G_CB(int n, IntPtr p_x,
	IpoptBoolType new_x, int m, int nele_jac,
	IntPtr p_iRow,
	IntPtr p_jCol,
	IntPtr p_values, IntPtr p_user_data);
#else
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	public delegate IpoptBoolType Eval_Jac_G_CB(int n, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] double[] x,
	IpoptBoolType new_x, int m, int nele_jac,
	[In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)] int[] iRow,
	[In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)] int[] jCol,
	[In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)] double[] values, IntPtr p_user_data);
#endif

	/// <summary>
	/// Delegate defining the callback function for evaluating the Hessian of
	/// the Lagrangian function.  Return value should be set to false if
	/// there was a problem doing the evaluation.
	/// </summary>
	/// <param name="n">Number of problem variables</param>
	/// <param name="x">Problem variables</param>
	/// <param name="new_x">true if problem variables are new for this call, false otherwise</param>
	/// <param name="obj_factor">Multiplier for the objective function in the Lagrangian</param>
	/// <param name="m">Number of constraint functions</param>
	/// <param name="lambda">Multipliers for the constraint functions in the Lagrangian</param>
	/// <param name="new_lambda">true if lambda values are new for this call, false otherwise</param>
	/// <param name="nele_hess">Number of non-zero elements in the Hessian</param>
	/// <param name="iRow">Row indices of the non-zero Hessian elements, defined here if values is null</param>
	/// <param name="jCol">Column indices of the non-zero Hessian elements, defined here if values is null</param>
	/// <param name="values">Values of the non-zero Hessian elements</param>
	/// <param name="p_user_data">Optional pointer to user defined data</param>
	/// <returns>true if evaluation succeeded, false otherwise</returns>
#if MONO
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	public delegate IpoptBoolType Eval_H_CB(int n, IntPtr p_x,
	IpoptBoolType new_x, double obj_factor, int m, IntPtr p_lambda,
	IpoptBoolType new_lambda, int nele_hess,
	IntPtr p_iRow,
	IntPtr p_jCol,
	IntPtr p_values, IntPtr p_user_data);
#else
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	public delegate IpoptBoolType Eval_H_CB(int n, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] double[] x,
	IpoptBoolType new_x, double obj_factor, int m, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)] double[] lambda,
	IpoptBoolType new_lambda, int nele_hess,
	[In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 7)] int[] iRow,
	[In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 7)] int[] jCol,
	[In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 7)] double[] values, IntPtr p_user_data);
#endif

	/// <summary>
	/// Delegate defining the callback function for giving intermediate
	/// execution control to the user.  If set, it is called once per
	/// iteration, providing the user with some information on the state
	/// of the optimization.  This can be used to print some
	/// user-defined output.  It also gives the user a way to terminate
	/// the optimization prematurely.  If this method returns false,
	/// Ipopt will terminate the optimization.
	/// </summary>
	/// <param name="alg_mod">Current Ipopt algorithm mode</param>
	/// <param name="iter_count">Current iteration count</param>
	/// <param name="obj_value">The unscaled objective value at the current point</param>
	/// <param name="inf_pr">The scaled primal infeasibility at the current point</param>
	/// <param name="inf_du">The scaled dual infeasibility at the current point</param>
	/// <param name="mu">The barrier parameter value at the current point</param>
	/// <param name="d_norm">The infinity norm (max) of the primal step 
	/// (for the original variables x and the internal slack variables s)</param>
	/// <param name="regularization_size">Value of the regularization term for the Hessian of the Lagrangian 
	/// in the augmented system</param>
	/// <param name="alpha_du">The stepsize for the dual variables</param>
	/// <param name="alpha_pr">The stepsize for the primal variables</param>
	/// <param name="ls_trials">The number of backtracking line search steps</param>
	/// <param name="p_user_data">Optional pointer to user defined data</param>
	/// <returns>true if the optimization should proceeded after callback return, false if
	/// the optimization should be terminated prematurely</returns>
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	public delegate IpoptBoolType Intermediate_CB(IpoptAlgorithmMode alg_mod, int iter_count, double obj_value, double inf_pr,
	double inf_du, double mu, double d_norm, double regularization_size, double alpha_du, double alpha_pr, int ls_trials, 
	IntPtr p_user_data);

	#endregion

	/// <summary>
	/// Static class providing the P/Invoke signatures to the Ipopt C interface.
	/// </summary>
	public static class IpoptAdapter
	{
		#region FIELDS

		private const string IpoptDllName = "IpOpt-vc10";

		#endregion

		#region P/INVOKE METHODS

		/// <summary>
		/// This function returns an object that can be passed to the IpoptSolve call.  It
		/// contains the basic definition of the optimization problem, such
		/// as number of variables and constraints, bounds on variables and
		/// constraints, information about the derivatives, and the callback
		/// function for the computation of the optimization problem
		/// functions and derivatives.  During this call, the options file
		/// ipopt.opt is read as well.
		/// </summary>
		/// <param name="n">Number of optimization variables</param>
		/// <param name="x_L">Lower bounds on variables. This array of size n is copied internally, so that the
		/// caller can change the incoming data after return without that IpoptProblem is modified.  Any value 
		/// less or equal than the number specified by option 'nlp_lower_bound_inf' is interpreted to be minus infinity.</param>
		/// <param name="x_U">Upper bounds on variables. This array of size n is copied internally, so that the
		/// caller can change the incoming data after return without that IpoptProblem is modified.  Any value 
		/// greater or equal than the number specified by option 'nlp_upper_bound_inf' is interpreted to be plus infinity.</param>
		/// <param name="m">Number of constraints.</param>
		/// <param name="g_L">Lower bounds on constraints. This array of size m is copied internally, so that the
		/// caller can change the incoming data after return without that IpoptProblem is modified.  Any value 
		/// less or equal than the number specified by option 'nlp_lower_bound_inf' is interpreted to be minus infinity.</param>
		/// <param name="g_U">Upper bounds on constraints. This array of size m is copied internally, so that the
		/// caller can change the incoming data after return without that IpoptProblem is modified.  Any value 
		/// greater or equal than the number specified by option 'nlp_upper_bound_inf' is interpreted to be plus infinity.</param>
		/// <param name="nele_jac">Number of non-zero elements in constraint Jacobian.</param>
		/// <param name="nele_hess">Number of non-zero elements in Hessian of Lagrangian.</param>
		/// <param name="index_style">Indexing style for iRow and jCol, 0 for C style, 1 for Fortran style</param>
		/// <param name="eval_f">Callback function for evaluating objective function</param>
		/// <param name="eval_g">Callback function for evaluating constraint functions</param>
		/// <param name="eval_grad_f">Callback function for evaluating gradient of objective function</param>
		/// <param name="eval_jac_g">Callback function for evaluating Jacobian of constraint functions</param>
		/// <param name="eval_h">Callback function for evaluating Hessian of Lagrangian function</param>
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern IntPtr CreateIpoptProblem(int n, double[] x_L, double[] x_U, int m, double[] g_L, double[] g_U, 
			int nele_jac, int nele_hess, IpoptIndexStyle index_style, 
			Eval_F_CB eval_f, Eval_G_CB eval_g, Eval_Grad_F_CB eval_grad_f, Eval_Jac_G_CB eval_jac_g, Eval_H_CB eval_h);

		/// <summary>
		/// Method for freeing a previously created IpoptProblem.  After freeing an IpoptProblem, it cannot be used anymore.
		/// </summary>
		/// <param name="ipopt_problem">Pointer to Ipopt problem.</param>
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern void FreeIpoptProblem(IntPtr ipopt_problem);

		/// <summary>
		/// Function for adding a string option.
		/// </summary>
		/// <param name="ipopt_problem">Pointer to Ipopt problem</param>
		/// <param name="keyword">Name of option</param>
		/// <param name="val">String value of option</param>
		/// <returns>true if setting option succeeded, false if the option could not be set (e.g., if keyword is unknown)</returns>
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern IpoptBoolType AddIpoptStrOption(IntPtr ipopt_problem, string keyword, string val);

		/// <summary>
		/// Function for adding a floating point option.
		/// </summary>
		/// <param name="ipopt_problem">Pointer to Ipopt problem</param>
		/// <param name="keyword">Name of option</param>
		/// <param name="val">Floating point value of option</param>
		/// <returns>true if setting option succeeded, false if the option could not be set (e.g., if keyword is unknown)</returns>
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern IpoptBoolType AddIpoptNumOption(IntPtr ipopt_problem, string keyword, double val);

		/// <summary>
		/// Function for adding an integer option.
		/// </summary>
		/// <param name="ipopt_problem">Pointer to Ipopt problem</param>
		/// <param name="keyword">Name of option</param>
		/// <param name="val">Integer value of option</param>
		/// <returns>true if setting option succeeded, false if the option could not be set (e.g., if keyword is unknown)</returns>
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern IpoptBoolType AddIpoptIntOption(IntPtr ipopt_problem, string keyword, int val);

#if !SILVERLIGHT
		/// <summary>
		/// Method for opening an output file for a given name with given printlevel.
		/// </summary>
		/// <param name="ipopt_problem">Pointer to Ipopt problem</param>
		/// <param name="file_name">Name of output file</param>
		/// <param name="print_level">Level of printed information</param>
		/// <returns>False, if there was a problem opening the file.</returns>
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern IpoptBoolType OpenIpoptOutputFile(IntPtr ipopt_problem, string file_name, int print_level);
#endif

		/// <summary>
		/// Optional function for setting scaling parameter for the NLP.
		/// This corresponds to the get_scaling_parameters method in TNLP.
		/// If the pointers x_scaling or g_scaling are null, then no scaling
		/// for x resp. g is done.
		/// </summary>
		/// <param name="ipopt_problem">Pointer to Ipopt problem</param>
		/// <param name="obj_scaling">Scaling of the objective function</param>
		/// <param name="x_scaling">Scaling of the problem variables</param>
		/// <param name="g_scaling">Scaling of the constraint functions</param>
		/// <returns>true if scaling succeeded, false otherwise</returns>
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern IpoptBoolType SetIpoptProblemScaling(
			IntPtr ipopt_problem, double obj_scaling, double[] x_scaling, double[] g_scaling);

		/// <summary>
		/// Setting a callback function for the "intermediate callback"
		/// method in the optimizer.  This gives control back to the user once
		/// per iteration.  If set, it provides the user with some
		/// information on the state of the optimization.  This can be used
		/// to print some user-defined output.  It also gives the user a way
		/// to terminate the optimization prematurely.  If the callback
		/// method returns false, Ipopt will terminate the optimization.
		/// Calling this set method to set the CB pointer to null disables
		/// the intermediate callback functionality.
		/// </summary>
		/// <param name="ipopt_problem">Pointer to Ipopt problem</param>
		/// <param name="intermediate_cb">Intermediate callback function</param>
		/// <returns>true if the callback function could be set successfully, false otherwise</returns>
#if !PRE39
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern IpoptBoolType SetIntermediateCallback(IntPtr ipopt_problem, Intermediate_CB intermediate_cb);
#else
		public static IpoptBoolType SetIntermediateCallback(IntPtr ipopt_problem, Intermediate_CB intermediate_cb)
		{
			return IpoptBoolType.False;
		}
#endif

		/// <summary>
		/// Function calling the IPOPT optimization algorithm for a problem previously defined with the constructor.
		/// IPOPT will use the options previously specified with AddOption for this problem.
		/// </summary>
		/// <param name="ipopt_problem">Pointer to Ipopt problem</param>
		/// <param name="x">Input: Starting point; Output: Optimal solution</param>
		/// <param name="obj_val">Final value of objective function (output only - ignored if null on input)</param>
		/// <param name="g">Values of constraint at final point (output only - ignored if null on input)</param>
		/// <param name="mult_g">Final multipliers for constraints (output only - ignored if null on input)</param>
		/// <param name="mult_x_L">Final multipliers for lower variable bounds (output only - ignored if null on input)</param>
		/// <param name="mult_x_U">Final multipliers for upper variable bounds (output only - ignored if null on input)</param>
		/// <param name="p_user_data">Optional pointer to user data</param>
		/// <returns>Outcome of the optimization procedure (e.g., success, failure etc).</returns>
		[DllImport(IpoptDllName, CallingConvention = CallingConvention.Cdecl)]
		public static extern IpoptReturnCode IpoptSolve(IntPtr ipopt_problem, double[] x, double[] g,
			out double obj_val, double[] mult_g, double[] mult_x_L, double[] mult_x_U, IntPtr p_user_data);

		#endregion
	}
}
