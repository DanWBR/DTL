// Copyright (C) 2010-2011 Anders Gustafsson and others. All Rights Reserved.
// This code is published under the Eclipse Public License.
//
// Author:  Anders Gustafsson, Cureos AB 2011-12-13

using System;
using System.Runtime.InteropServices;

namespace Cureos.Numerics
{
    #region PUBLIC DELEGATES

    /// <summary>
    /// Delegate defining the callback function for evaluating the value of
    /// the objective function.  Return value should be set to false if
    /// there was a problem doing the evaluation.
    /// </summary>
    /// <param name="n">Number of problem variables</param>
    /// <param name="x">Problem variables</param>
    /// <param name="new_x">true if problem variables are new for this call, false otherwise</param>
    /// <param name="obj_value">Evaluated objective function value</param>
    /// <returns>true if evaluation succeeded, false otherwise</returns>
    public delegate bool EvaluateObjectiveDelegate(int n, double[] x, bool new_x, out double obj_value);

    /// <summary>
    /// Delegate defining the callback function for evaluating the gradient of
    /// the objective function.  Return value should be set to false if
    /// there was a problem doing the evaluation.
    /// </summary>
    /// <param name="n">Number of problem variables</param>
    /// <param name="x">Problem variables</param>
    /// <param name="new_x">true if problem variables are new for this call, false otherwise</param>
    /// <param name="grad_f">Objective function gradient</param>
    /// <returns>true if evaluation succeeded, false otherwise</returns>
    public delegate bool EvaluateObjectiveGradientDelegate(int n, double[] x, bool new_x, double[] grad_f);

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
    /// <returns>true if evaluation succeeded, false otherwise</returns>
    public delegate bool EvaluateConstraintsDelegate(int n, double[] x, bool new_x, int m, double[] g);

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
    /// <returns>true if evaluation succeeded, false otherwise</returns>
    public delegate bool EvaluateJacobianDelegate(int n, double[] x, bool new_x, int m, int nele_jac, int[] iRow, int[] jCol, double[] values);

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
    /// <returns>true if evaluation succeeded, false otherwise</returns>
    public delegate bool EvaluateHessianDelegate(int n, double[] x, bool new_x, double obj_factor, int m, double[] lambda, bool new_lambda, 
    int nele_hess, int[] iRow, int[] jCol, double[] values);

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
    /// <returns>true if the optimization should proceeded after callback return, false if
    /// the optimization should be terminated prematurely</returns>
    public delegate bool IntermediateDelegate(IpoptAlgorithmMode alg_mod, int iter_count, double obj_value, double inf_pr, double inf_du,
    double mu, double d_norm, double regularization_size, double alpha_du, double alpha_pr, int ls_trials);

    #endregion

    /// <summary>
    /// Managed wrapper and base class for setup and solution of an Ipopt problem.
    /// </summary>
    public class IpoptProblem : IDisposable
    {
        #region INTERNAL CALLBACK FUNCTION CLASSES

        private class ObjectiveEvaluator
        {
            private readonly EvaluateObjectiveDelegate m_eval_f_cb;

            internal ObjectiveEvaluator(EvaluateObjectiveDelegate eval_f_cb)
            {
                m_eval_f_cb = eval_f_cb;
            }

#if SILVERLIGHT
            [AllowReversePInvokeCalls]
#endif
            internal IpoptBoolType Evaluate(int n, double[] x, IpoptBoolType new_x, out double obj_value, IntPtr p_user_data)
            {
                return m_eval_f_cb(n, x, new_x == IpoptBoolType.True, out obj_value) ? IpoptBoolType.True : IpoptBoolType.False;
            }
        }

        private class ConstraintsEvaluator
        {
            private readonly EvaluateConstraintsDelegate m_eval_g_cb;

            internal ConstraintsEvaluator(EvaluateConstraintsDelegate eval_g_cb)
            {
                m_eval_g_cb = eval_g_cb;
            }

#if SILVERLIGHT
            [AllowReversePInvokeCalls]
#endif
            internal IpoptBoolType Evaluate(int n, double[] x, IpoptBoolType new_x, int m, double[] g, IntPtr p_user_data)
            {
                return m_eval_g_cb(n, x, new_x == IpoptBoolType.True, m, g) ? IpoptBoolType.True : IpoptBoolType.False;
            }
        }

        private class ObjectiveGradientEvaluator
        {
            private readonly EvaluateObjectiveGradientDelegate m_eval_grad_f_cb;

            internal ObjectiveGradientEvaluator(EvaluateObjectiveGradientDelegate eval_grad_f_cb)
            {
                m_eval_grad_f_cb = eval_grad_f_cb;
            }

#if SILVERLIGHT
            [AllowReversePInvokeCalls]
#endif
            internal IpoptBoolType Evaluate(int n, double[] x, IpoptBoolType new_x, double[] grad_f, IntPtr p_user_data)
            {
                return m_eval_grad_f_cb(n, x, new_x == IpoptBoolType.True, grad_f) ? IpoptBoolType.True : IpoptBoolType.False;
            }
        }

        private class JacobianEvaluator
        {
            private readonly EvaluateJacobianDelegate m_eval_jac_g_cb;

            internal JacobianEvaluator(EvaluateJacobianDelegate eval_jac_g_cb)
            {
                m_eval_jac_g_cb = eval_jac_g_cb;
            }

#if MONO
            [AllowReversePInvokeCalls]
            internal IpoptBoolType Evaluate(int n, IntPtr p_x, IpoptBoolType new_x, int m, int nele_jac,
                IntPtr p_iRow, IntPtr p_jCol, IntPtr p_values, IntPtr p_user_data)
            {
                var x = new double[n];
                var iRow = new int[nele_jac];
                var jCol = new int[nele_jac];

                if (p_x != IntPtr.Zero) Marshal.Copy(p_x, x, 0, n);
                if (p_iRow != IntPtr.Zero) Marshal.Copy(p_iRow, iRow, 0, nele_jac);
                if (p_jCol != IntPtr.Zero) Marshal.Copy(p_jCol, jCol, 0, nele_jac);
                var values = p_values != IntPtr.Zero ? new double[nele_jac] : null;

                var ret = m_eval_jac_g_cb(n, x, new_x == IpoptBoolType.True, m, nele_jac, iRow, jCol, values);

                if (p_values != IntPtr.Zero) Marshal.Copy(values, 0, p_values, nele_jac);
                if (p_iRow != IntPtr.Zero) Marshal.Copy(iRow, 0, p_iRow, nele_jac);
                if (p_jCol != IntPtr.Zero) Marshal.Copy(jCol, 0, p_jCol, nele_jac);

                return ret ? IpoptBoolType.True : IpoptBoolType.False;
            }
#else
#if SILVERLIGHT
            [AllowReversePInvokeCalls]
#endif
            internal IpoptBoolType Evaluate(int n, double[] x, IpoptBoolType new_x, int m, int nele_jac, 
                int[] iRow, int[] jCol, double[] values, IntPtr p_user_data)
            {
                return m_eval_jac_g_cb(n, x, new_x == IpoptBoolType.True, m, nele_jac, iRow, jCol, values)
                           ? IpoptBoolType.True
                           : IpoptBoolType.False;
            }
#endif
        }

        private class HessianEvaluator
        {
            private readonly EvaluateHessianDelegate m_eval_h_cb;

            internal HessianEvaluator(EvaluateHessianDelegate eval_h_cb)
            {
                m_eval_h_cb = eval_h_cb;
            }

#if MONO
            [AllowReversePInvokeCalls]
            internal IpoptBoolType Evaluate(int n, IntPtr p_x, IpoptBoolType new_x, double obj_factor, int m, IntPtr p_lambda,
                IpoptBoolType new_lambda, int nele_hess, IntPtr p_iRow, IntPtr p_jCol, IntPtr p_values, IntPtr p_user_data)
            {
                var x = new double[n];
                var lambda = new double[m];
                var iRow = new int[nele_hess];
                var jCol = new int[nele_hess];

                if (p_x != IntPtr.Zero) Marshal.Copy(p_x, x, 0, n);
                if (p_lambda != IntPtr.Zero) Marshal.Copy(p_lambda, lambda, 0, m);
                if (p_iRow != IntPtr.Zero) Marshal.Copy(p_iRow, iRow, 0, nele_hess);
                if (p_jCol != IntPtr.Zero) Marshal.Copy(p_jCol, jCol, 0, nele_hess);
                var values = p_values != IntPtr.Zero ? new double[nele_hess] : null;

                var ret = m_eval_h_cb(n, x, new_x == IpoptBoolType.True, obj_factor, m, lambda,
                                new_lambda == IpoptBoolType.True, nele_hess, iRow, jCol, values);

                if (p_values != IntPtr.Zero) Marshal.Copy(values, 0, p_values, nele_hess);
                if (p_iRow != IntPtr.Zero) Marshal.Copy(iRow, 0, p_iRow, nele_hess);
                if (p_jCol != IntPtr.Zero) Marshal.Copy(jCol, 0, p_jCol, nele_hess);

                return ret ? IpoptBoolType.True : IpoptBoolType.False;
            }
#else
#if SILVERLIGHT
            [AllowReversePInvokeCalls]
#endif
            internal IpoptBoolType Evaluate(int n, double[] x, IpoptBoolType new_x, double obj_factor, int m, double[] lambda,
                IpoptBoolType new_lambda, int nele_hess, int[] iRow, int[] jCol, double[] values, IntPtr p_user_data)
            {
                return m_eval_h_cb(n, x, new_x == IpoptBoolType.True, obj_factor, m, lambda,
                                new_lambda == IpoptBoolType.True, nele_hess, iRow, jCol, values)
                           ? IpoptBoolType.True
                           : IpoptBoolType.False;
            }
#endif
        }

        private class IntermediateReporter
        {
            private readonly IntermediateDelegate m_intermediate_cb;

            internal IntermediateReporter(IntermediateDelegate intermediate_cb)
            {
                m_intermediate_cb = intermediate_cb;
            }

#if SILVERLIGHT
            [AllowReversePInvokeCalls]
#endif
            internal IpoptBoolType Report(IpoptAlgorithmMode alg_mod, int iter_count, double obj_value, double inf_pr, double inf_du,
                double mu, double d_norm, double regularization_size, double alpha_du, double alpha_pr, int ls_trials, IntPtr p_user_data)
            {
                return m_intermediate_cb(alg_mod, iter_count, obj_value, inf_pr, inf_du, mu, d_norm, 
                    regularization_size, alpha_du, alpha_pr, ls_trials) ? IpoptBoolType.True : IpoptBoolType.False;
            }
        }

        #endregion

        #region FIELDS

        /// <summary>
        /// Value to indicate that a variable or constraint function has no upper bound 
        /// (provided that IPOPT option "nlp_upper_bound_inf" is less than 2e19)
        /// </summary>
        public const double PositiveInfinity = 2.0e19;

        /// <summary>
        /// Value to indicate that a variable or constraint function has no lower bound 
        /// (provided that IPOPT option "nlp_lower_bound_inf" is greater than -2e19)
        /// </summary>
        public const double NegativeInfinity = -2.0e19;

        private IntPtr m_problem;
        private bool m_disposed;

        private readonly Eval_F_CB m_eval_f_cb;
        private readonly Eval_G_CB m_eval_g_cb; 
        private readonly Eval_Grad_F_CB m_eval_grad_f_cb;
        private readonly Eval_Jac_G_CB m_eval_jac_g_cb;
        private readonly Eval_H_CB m_eval_h_cb;
        private Intermediate_CB m_intermediate_cb;

        #endregion

        #region CONSTRUCTORS

        /// <summary>
        /// Constructor for creating a new Ipopt Problem object using managed 
        /// function delegates.  This function
        /// initializes an object that can be passed to the IpoptSolve call.  It
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
        /// <param name="eval_f_cb">Managed callback function for evaluating objective function</param>
        /// <param name="eval_g_cb">Managed callback function for evaluating constraint functions</param>
        /// <param name="eval_grad_f_cb">Managed callback function for evaluating gradient of objective function</param>
        /// <param name="eval_jac_g_cb">Managed callback function for evaluating Jacobian of constraint functions</param>
        /// <param name="eval_h_cb">Managed callback function for evaluating Hessian of Lagrangian function</param>
        public IpoptProblem(int n, double[] x_L, double[] x_U, int m, double[] g_L, double[] g_U, int nele_jac, int nele_hess,
            EvaluateObjectiveDelegate eval_f_cb, EvaluateConstraintsDelegate eval_g_cb, EvaluateObjectiveGradientDelegate eval_grad_f_cb,
            EvaluateJacobianDelegate eval_jac_g_cb, EvaluateHessianDelegate eval_h_cb)
        {
            m_eval_f_cb = new ObjectiveEvaluator(eval_f_cb).Evaluate;
            m_eval_g_cb = new ConstraintsEvaluator(eval_g_cb).Evaluate;
            m_eval_grad_f_cb = new ObjectiveGradientEvaluator(eval_grad_f_cb).Evaluate;
            m_eval_jac_g_cb = new JacobianEvaluator(eval_jac_g_cb).Evaluate;
            m_eval_h_cb = new HessianEvaluator(eval_h_cb).Evaluate;
            m_intermediate_cb = null;

            m_problem = IpoptAdapter.CreateIpoptProblem(n, x_L, x_U, m, g_L, g_U, nele_jac, nele_hess, IpoptIndexStyle.C,
                m_eval_f_cb, m_eval_g_cb, m_eval_grad_f_cb, m_eval_jac_g_cb, m_eval_h_cb);

            m_disposed = false;
        }

        /// <summary>
        /// Constructor for creating a new Ipopt Problem object using native
        /// function delegates. This function
        /// initializes an object that can be passed to the IpoptSolve call.  It
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
        /// <param name="eval_f_cb">Native callback function for evaluating objective function</param>
        /// <param name="eval_g_cb">Native callback function for evaluating constraint functions</param>
        /// <param name="eval_grad_f_cb">Native callback function for evaluating gradient of objective function</param>
        /// <param name="eval_jac_g_cb">Native callback function for evaluating Jacobian of constraint functions</param>
        /// <param name="eval_h_cb">Native callback function for evaluating Hessian of Lagrangian function</param>
        public IpoptProblem(int n, double[] x_L, double[] x_U, int m, double[] g_L, double[] g_U, int nele_jac, int nele_hess,
            Eval_F_CB eval_f_cb, Eval_G_CB eval_g_cb, Eval_Grad_F_CB eval_grad_f_cb, Eval_Jac_G_CB eval_jac_g_cb, Eval_H_CB eval_h_cb)
        {
            m_eval_f_cb = eval_f_cb;
            m_eval_g_cb = eval_g_cb;
            m_eval_grad_f_cb = eval_grad_f_cb;
            m_eval_jac_g_cb = eval_jac_g_cb;
            m_eval_h_cb = eval_h_cb;
            m_intermediate_cb = null;

            m_problem = IpoptAdapter.CreateIpoptProblem(n, x_L, x_U, m, g_L, g_U, nele_jac, nele_hess, IpoptIndexStyle.C,
                m_eval_f_cb, m_eval_g_cb, m_eval_grad_f_cb, m_eval_jac_g_cb, m_eval_h_cb);
            
            m_disposed = false;
        }

        /// <summary>
        /// Constructor for creating a subclassed Ipopt Problem object using managed or
        /// native function delegates. This is the preferred constructor when
        /// subclassing IpoptProblem. Prerequisite is that the managed/native optimization 
        /// function delegates are implemented in the inheriting class.
        /// This function
        /// initializes an object that can be passed to the IpoptSolve call.  It
        /// contains the basic definition of the optimization problem, such
        /// as number of variables and constraints, bounds on variables and
        /// constraints, information about the derivatives, and the callback
        /// function for the computation of the optimization problem
        /// functions and derivatives. During this call, the options file
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
        /// <param name="useNativeCallbackFunctions">If set to true, native callback functions are used to setup
        /// the Ipopt problem; if set to false, managed callback functions are used.</param>
        /// <param name="useHessianApproximation">If set to true, the Ipopt optimizer creates a limited memory
        /// Hessian approximation and the eval_h (managed or native) method need not be implemented. 
        /// If set to false, an exact Hessian should be evaluated using the appropriate Hessian evaluation method.</param>
        /// <param name="useIntermediateCallback">If set to true, the intermediate method (managed or native) will be called 
        /// after each full iteration. If false, the intermediate callback function will not be called.</param>
        protected IpoptProblem(int n, double[] x_L, double[] x_U, int m, double[] g_L, double[] g_U, int nele_jac, int nele_hess,
            bool useNativeCallbackFunctions = false, bool useHessianApproximation = false, bool useIntermediateCallback = false)
        {
            if (useNativeCallbackFunctions)
            {
                m_eval_f_cb = eval_f;
                m_eval_g_cb = eval_g;
                m_eval_grad_f_cb = eval_grad_f;
                m_eval_jac_g_cb = eval_jac_g;
                m_eval_h_cb = eval_h;
            }
            else
            {
                m_eval_f_cb = new ObjectiveEvaluator(eval_f).Evaluate;
                m_eval_g_cb = new ConstraintsEvaluator(eval_g).Evaluate;
                m_eval_grad_f_cb = new ObjectiveGradientEvaluator(eval_grad_f).Evaluate;
                m_eval_jac_g_cb = new JacobianEvaluator(eval_jac_g).Evaluate;
                m_eval_h_cb = new HessianEvaluator(eval_h).Evaluate;
            }
            m_intermediate_cb = null;

            m_problem = IpoptAdapter.CreateIpoptProblem(n, x_L, x_U, m, g_L, g_U, nele_jac, nele_hess, IpoptIndexStyle.C,
                                           m_eval_f_cb, m_eval_g_cb, m_eval_grad_f_cb, m_eval_jac_g_cb, m_eval_h_cb);

            if (useHessianApproximation) AddOption("hessian_approximation", "limited-memory");

            if (useIntermediateCallback)
            {
                if (useNativeCallbackFunctions)
                    SetIntermediateCallback((Intermediate_CB)intermediate);
                else
                    SetIntermediateCallback((IntermediateDelegate)intermediate);
            }

            m_disposed = false;
        }

        /// <summary>
        /// Destructor for IPOPT problem
        /// </summary>
        ~IpoptProblem()
        {
            Dispose(false);
        }

        #endregion

        #region PROPERTIES

        /// <summary>
        /// Gets the initialization status of the Ipopt problem.
        /// </summary>
        public bool IsInitialized
        {
            get { return m_problem != IntPtr.Zero; }
        }

        #endregion

        #region METHODS

        /// <summary>
        /// Implement the IDisposable interface to release the Ipopt DLL resources held by this class
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Function for adding a string option.
        /// </summary>
        /// <param name="keyword">Name of option</param>
        /// <param name="val">String value of option</param>
        /// <returns>true if setting option succeeded, false if the option could not be set (e.g., if keyword is unknown)</returns>
        public bool AddOption(string keyword, string val)
        {
            return IsInitialized && IpoptAdapter.AddIpoptStrOption(m_problem, keyword, val) == IpoptBoolType.True;
        }

        /// <summary>
        /// Function for adding a floating point option.
        /// </summary>
        /// <param name="keyword">Name of option</param>
        /// <param name="val">Floating point value of option</param>
        /// <returns>true if setting option succeeded, false if the option could not be set (e.g., if keyword is unknown)</returns>
        public bool AddOption(string keyword, double val)
        {
            return IsInitialized && IpoptAdapter.AddIpoptNumOption(m_problem, keyword, val) == IpoptBoolType.True;
        }

        /// <summary>
        /// Function for adding an integer option.
        /// </summary>
        /// <param name="keyword">Name of option</param>
        /// <param name="val">Integer value of option</param>
        /// <returns>true if setting option succeeded, false if the option could not be set (e.g., if keyword is unknown)</returns>
        public bool AddOption(string keyword, int val)
        {
            return IsInitialized && IpoptAdapter.AddIpoptIntOption(m_problem, keyword, val) == IpoptBoolType.True;
        }

#if !SILVERLIGHT
        /// <summary>
        /// Method for opening an output file for a given name with given print level.
        /// </summary>
        /// <param name="file_name">Name of output file</param>
        /// <param name="print_level">Level of printed information</param>
        /// <returns>False, if there was a problem opening the file.</returns>
        public bool OpenOutputFile(string file_name, int print_level)
        {
            return IsInitialized && IpoptAdapter.OpenIpoptOutputFile(m_problem, file_name, print_level) == IpoptBoolType.True;
        }
#endif

        /// <summary>
        /// Optional function for setting scaling parameter for the NLP.
        /// This corresponds to the get_scaling_parameters method in TNLP.
        /// If the pointers x_scaling or g_scaling are null, then no scaling
        /// for x resp. g is done.
        /// </summary>
        /// <param name="obj_scaling">Scaling of the objective function</param>
        /// <param name="x_scaling">Scaling of the problem variables</param>
        /// <param name="g_scaling">Scaling of the constraint functions</param>
        /// <returns>true if scaling succeeded, false otherwise</returns>
        public bool SetScaling(double obj_scaling, double[] x_scaling, double[] g_scaling)
        {
            return IsInitialized &&
                   AddOption("nlp_scaling_method", "user-scaling") &&
                   IpoptAdapter.SetIpoptProblemScaling(m_problem, obj_scaling, x_scaling, g_scaling) ==
                   IpoptBoolType.True;
        }

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
        /// <param name="intermediate_cb">Managed intermediate callback function</param>
        /// <returns>true if the callback function could be set successfully, false otherwise</returns>
        public bool SetIntermediateCallback(IntermediateDelegate intermediate_cb)
        {
            return IsInitialized && IpoptAdapter.SetIntermediateCallback(
                m_problem, m_intermediate_cb = new IntermediateReporter(intermediate_cb).Report) == IpoptBoolType.True;
        }

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
        /// <param name="intermediate_cb">Native intermediate callback function</param>
        /// <returns>true if the callback function could be set successfully, false otherwise</returns>
        public bool SetIntermediateCallback(Intermediate_CB intermediate_cb)
        {
            return IsInitialized && 
                IpoptAdapter.SetIntermediateCallback(m_problem, m_intermediate_cb = intermediate_cb) == IpoptBoolType.True;
        }

        /// <summary>
        /// Function calling the IPOPT optimization algorithm for a problem previously defined with the constructor.
        /// IPOPT will use the options previously specified with AddOption for this problem.
        /// </summary>
        /// <param name="x">Input: Starting point; Output: Optimal solution</param>
        /// <param name="obj_val">Final value of objective function (output only)</param>
        /// <param name="g">Values of constraint at final point (output only - ignored if null on input)</param>
        /// <param name="mult_g">Final multipliers for constraints (output only - ignored if null on input)</param>
        /// <param name="mult_x_L">Final multipliers for lower variable bounds (output only - ignored if null on input)</param>
        /// <param name="mult_x_U">Final multipliers for upper variable bounds (output only - ignored if null on input)</param>
        /// <returns>Outcome of the optimization procedure (e.g., success, failure etc).</returns>
#if !NET20
        public IpoptReturnCode SolveProblem(double[] x, out double obj_val, double[] g = null, double[] mult_g = null, 
            double[] mult_x_L = null, double[] mult_x_U = null)
#else
        public IpoptReturnCode SolveProblem(double[] x, out double obj_val, double[] g, double[] mult_g, 
            double[] mult_x_L, double[] mult_x_U)
#endif
        {
            if (!IsInitialized)
            {
                obj_val = PositiveInfinity;
                return IpoptReturnCode.Problem_Not_Initialized;
            }
            return IpoptAdapter.IpoptSolve(m_problem, x, g, out obj_val, mult_g, mult_x_L, mult_x_U, IntPtr.Zero);
        }

        /// <summary>
        /// Dispose(bool disposing) executes in two distinct scenarios.
        /// If disposing equals true, the method has been called directly
        /// or indirectly by a user's code. Managed and unmanaged resources
        /// can be disposed.
        /// If disposing equals false, the method has been called by the
        /// runtime from inside the finalizer and you should not reference
        /// other objects. Only unmanaged resources can be disposed.
        /// </summary>
        /// <param name="disposing">true if Dispose method is explicitly called, false otherwise.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!m_disposed)
            {
                if (m_problem != IntPtr.Zero)
                    IpoptAdapter.FreeIpoptProblem(m_problem);

                if (disposing)
                {
                    m_problem = IntPtr.Zero;
                }

                m_disposed = true;
            }
        }

        #endregion

        #region BASE CLASS NATIVE CALLBACK IMPLEMENTATIONS

        /// <summary>
        /// Dummy implementation of the native objective function evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
#if SILVERLIGHT
        [AllowReversePInvokeCalls]
#endif
        public virtual IpoptBoolType eval_f(int n, double[] x, IpoptBoolType new_x, out double obj_value, IntPtr p_user_data)
        {
            throw new NotSupportedException("Objective function evaluation should be implemented in subclass.");
        }

        /// <summary>
        /// Dummy implementation of the native constraints evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
#if SILVERLIGHT
        [AllowReversePInvokeCalls]
#endif
        public virtual IpoptBoolType eval_g(int n, double[] x, IpoptBoolType new_x, int m, double[] g, IntPtr p_user_data)
        {
            throw new NotSupportedException("Constraints evaluation should be implemented in subclass.");
        }

        /// <summary>
        /// Dummy implementation of the native objective function gradient evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
#if SILVERLIGHT
        [AllowReversePInvokeCalls]
#endif
        public virtual IpoptBoolType eval_grad_f(int n, double[] x, IpoptBoolType new_x, double[] grad_f, IntPtr p_user_data)
        {
            throw new NotSupportedException("Objective function gradient evaluation should be implemented in subclass.");
        }

        /// <summary>
        /// Dummy implementation of the native Jacobian evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
#if MONO
        [AllowReversePInvokeCalls]
        public virtual IpoptBoolType eval_jac_g(int n, IntPtr p_x, IpoptBoolType new_x, int m, int nele_jac,
            IntPtr p_iRow, IntPtr p_jCol, IntPtr p_values, IntPtr p_user_data)
        {
            throw new NotSupportedException("Jacobian evaluation should be implemented in subclass.");
        }
#else
#if SILVERLIGHT
        [AllowReversePInvokeCalls]
#endif
        public virtual IpoptBoolType eval_jac_g(int n, double[] x, IpoptBoolType new_x, int m, int nele_jac,
            int[] iRow, int[] jCol, double[] values, IntPtr p_user_data)
        {
            throw new NotSupportedException("Jacobian evaluation should be implemented in subclass.");
        }
#endif

        /// <summary>
        /// Dummy implementation of the native Hessian evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
#if MONO
        [AllowReversePInvokeCalls]
        public virtual IpoptBoolType eval_h(int n, IntPtr p_x, IpoptBoolType new_x, double obj_factor, int m, IntPtr p_lambda,
            IpoptBoolType new_lambda, int nele_hess, IntPtr p_iRow, IntPtr p_jCol, IntPtr p_values, IntPtr p_user_data)
        {
            throw new NotSupportedException("Hessian evaluation should be implemented in subclass.");
        }
#else
#if SILVERLIGHT
        [AllowReversePInvokeCalls]
#endif
        public virtual IpoptBoolType eval_h(int n, double[] x, IpoptBoolType new_x, double obj_factor, int m, double[] lambda,
            IpoptBoolType new_lambda, int nele_hess, int[] iRow, int[] jCol, double[] values, IntPtr p_user_data)
        {
            throw new NotSupportedException("Hessian evaluation should be implemented in subclass.");
        }
#endif

        /// <summary>
        /// Dummy implementation of the native intermediate callback delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
#if SILVERLIGHT
        [AllowReversePInvokeCalls]
#endif
        public virtual IpoptBoolType intermediate(IpoptAlgorithmMode alg_mod, int iter_count, double obj_value, double inf_pr, double inf_du,
            double mu, double d_norm, double regularization_size, double alpha_du, double alpha_pr, int ls_trials, IntPtr p_user_data)
        {
            throw new NotSupportedException("Intermediate callback method should be implemented in subclass.");
        }

        #endregion

        #region BASE CLASS OVERRIDABLE MANAGED CALLBACK IMPLEMENTATIONS

        /// <summary>
        /// Dummy implementation of the managed objective function evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
        public virtual bool eval_f(int n, double[] x, bool new_x, out double obj_value)
        {
            throw new NotSupportedException("Objective function evaluation should be implemented in subclass.");
        }

        /// <summary>
        /// Dummy implementation of the managed constraints evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
        public virtual bool eval_g(int n, double[] x, bool new_x, int m, double[] g)
        {
            throw new NotSupportedException("Contraints evaluation should be implemented in subclass.");
        }

        /// <summary>
        /// Dummy implementation of the managed objective function gradient evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
        public virtual bool eval_grad_f(int n, double[] x, bool new_x, double[] grad_f)
        {
            throw new NotSupportedException("Objective function gradient evaluation should be implemented in subclass.");
        }

        /// <summary>
        /// Dummy implementation of the managed Jacobian evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
        public virtual bool eval_jac_g(int n, double[] x, bool new_x, int m, int nele_jac, int[] iRow, int[] jCol, double[] values)
        {
            throw new NotSupportedException("Jacobian evaluation should be implemented in subclass.");
        }

        /// <summary>
        /// Dummy implementation of the managed Hessian evaluation delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
        public virtual bool eval_h(int n, double[] x, bool new_x, double obj_factor, int m, double[] lambda, bool new_lambda,
                    int nele_hess, int[] iRow, int[] jCol, double[] values)
        {
            throw new NotSupportedException("Hessian evaluation should be implemented in subclass.");
        }

        /// <summary>
        /// Dummy implementation of the managed intermediate callback method delegate.
        /// Recommended action is to override this method in a subclass and use
        /// the <see cref="IpoptProblem(int,double[],double[],int,double[],double[],int,int,bool,bool,bool)">protected constructor</see>
        /// of this base class to initialize the subclassed object.
        /// </summary>
        public virtual bool intermediate(IpoptAlgorithmMode alg_mod, int iter_count, double obj_value, double inf_pr, double inf_du,
            double mu, double d_norm, double regularization_size, double alpha_du, double alpha_pr, int ls_trials)
        {
            throw new NotSupportedException("Intermediate callback method should be implemented in subclass.");
        }

        #endregion
    }
}
