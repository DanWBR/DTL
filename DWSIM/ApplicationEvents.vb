Imports Cudafy

Namespace My

    ' The following events are availble for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication

        Public Shared _ResourceManager As System.Resources.ResourceManager
        Public Shared _HelpManager As System.Resources.ResourceManager
        Public Shared _PropertyNameManager As System.Resources.ResourceManager
        Public Shared _CultureInfo As System.Globalization.CultureInfo
        Public Shared CalculatorStopRequested As Boolean = False
        Public Shared CommandLineMode As Boolean = False
        Public CAPEOPENMode As Boolean = False
        Public Shared IsRunningParallelTasks As Boolean = False

        Public Shared _EnableGPUProcessing As Boolean = False
        Public Shared _EnableParallelProcessing As Boolean = False
        Public Shared _DebugLevel As Integer = 0

        Public Shared gpu As Cudafy.Host.GPGPU
        Public Shared gpumod As CudafyModule

    End Class

End Namespace

