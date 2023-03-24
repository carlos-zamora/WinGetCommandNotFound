using System.Security.Principal;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Management.Automation.Subsystem;
using System.Management.Automation.Subsystem.Feedback;
using System.Management.Automation.Subsystem.Prediction;
using System.Runtime.InteropServices;
using Microsoft.Management.Deployment;

namespace wingetprovider
{
    // Adapted from https://github.com/microsoft/winget-cli/blob/1898da0b657585d2e6399ef783ecb667eed280f9/src/PowerShell/Microsoft.WinGet.Client/Helpers/ComObjectFactory.cs
    public class ComObjectFactory
    {
        private static readonly Guid PackageManagerClsid = Guid.Parse("C53A4F16-787E-42A4-B304-29EFFB4BF597");
        private static readonly Guid FindPackagesOptionsClsid = Guid.Parse("572DED96-9C60-4526-8F92-EE7D91D38C1A");
        private static readonly Guid PackageMatchFilterClsid = Guid.Parse("D02C9DAF-99DC-429C-B503-4E504E4AB000");

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "COM only usage.")]
        private static readonly Type PackageManagerType = Type.GetTypeFromCLSID(PackageManagerClsid);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "COM only usage.")]
        private static readonly Type FindPackagesOptionsType = Type.GetTypeFromCLSID(FindPackagesOptionsClsid);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "COM only usage.")]
        private static readonly Type PackageMatchFilterType = Type.GetTypeFromCLSID(PackageMatchFilterClsid);

        private static readonly Guid PackageManagerIid = Guid.Parse("B375E3B9-F2E0-5C93-87A7-B67497F7E593");
        private static readonly Guid FindPackagesOptionsIid = Guid.Parse("A5270EDD-7DA7-57A3-BACE-F2593553561F");
        private static readonly Guid PackageMatchFilterIid = Guid.Parse("D981ECA3-4DE5-5AD7-967A-698C7D60FC3B");

        public static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "COM only usage.")]
        private static T Create<T>(Type type, in Guid iid)
        {
            object instance = null;
            if (IsAdministrator())
            {
                var hr = WinGetServerManualActivation_CreateInstance(type.GUID, iid, 0, out instance);
                if (hr < 0)
                {
                    throw new COMException($"Failed to create instance: {hr}", hr);
                }
            }
            else
            {
                instance = Activator.CreateInstance(type);
            }

            IntPtr pointer = Marshal.GetIUnknownForObject(instance);
            return WinRT.MarshalInterface<T>.FromAbi(pointer);
        }

        [DllImport("winrtact.dll", EntryPoint = "WinGetServerManualActivation_CreateInstance", ExactSpelling = true, PreserveSig = true)]
        private static extern int WinGetServerManualActivation_CreateInstance(
                [In, MarshalAs(UnmanagedType.LPStruct)] Guid clsid,
                [In, MarshalAs(UnmanagedType.LPStruct)] Guid iid,
                uint flags,
                [Out, MarshalAs(UnmanagedType.IUnknown)] out object instance);

        [DllImport("winrtact.dll", EntryPoint = "winrtact_Initialize", ExactSpelling = true, PreserveSig = true)]
        public static extern void InitializeUndockedRegFreeWinRT();

        public static PackageManager CreatePackageManager()
        {
            return Create<PackageManager>(PackageManagerType, PackageManagerIid);
        }

        public static FindPackagesOptions CreateFindPackagesOptions()
        {
            return Create<FindPackagesOptions>(FindPackagesOptionsType, FindPackagesOptionsIid);
        }

        public static PackageMatchFilter CreatePackageMatchFilter()
        {
            return Create<PackageMatchFilter>(PackageMatchFilterType, PackageMatchFilterIid);
        }
    }

    public sealed class Init : IModuleAssemblyInitializer, IModuleAssemblyCleanup
    {
        internal const string id = "e5351aa4-dfde-4d4d-bf0f-1a2f5a37d8d6";

        public void OnImport()
        {
            if (!Platform.IsWindows)
            {
                return;
            }

            // Ensure WinGet is installed
            using (var rs = RunspaceFactory.CreateRunspace(InitialSessionState.CreateDefault()))
            {
                rs.Open();
                var invocation = rs.SessionStateProxy.InvokeCommand;
                var winget = invocation.GetCommand("winget", CommandTypes.Application);
                if (winget is null)
                {
                    return;
                }
            }

            SubsystemManager.RegisterSubsystem<IFeedbackProvider, WinGetCommandNotFoundFeedbackPredictor>(WinGetCommandNotFoundFeedbackPredictor.Singleton);
            SubsystemManager.RegisterSubsystem<ICommandPredictor, WinGetCommandNotFoundFeedbackPredictor>(WinGetCommandNotFoundFeedbackPredictor.Singleton);
        }

        public void OnRemove(PSModuleInfo psModuleInfo)
        {
            SubsystemManager.UnregisterSubsystem<IFeedbackProvider>(new Guid(id));
            SubsystemManager.UnregisterSubsystem<ICommandPredictor>(new Guid(id));
        }
    }

    public sealed class WinGetCommandNotFoundFeedbackPredictor : IFeedbackProvider, ICommandPredictor
    {
        private readonly Guid _guid;
        private string? _suggestion;
        private PackageManager _packageManager;
        private FindPackagesOptions _findPackagesOptions;
        private PackageMatchFilter _packageMatchFilter;
        Dictionary<string, string>? ISubsystem.FunctionsToDefine => null;

        public static WinGetCommandNotFoundFeedbackPredictor Singleton { get; } = new WinGetCommandNotFoundFeedbackPredictor(Init.id);
        private WinGetCommandNotFoundFeedbackPredictor(string guid)
        {
            _guid = new Guid(guid);
            ComObjectFactory.InitializeUndockedRegFreeWinRT();
            _packageManager = ComObjectFactory.CreatePackageManager();
            _findPackagesOptions = ComObjectFactory.CreateFindPackagesOptions();
            _packageMatchFilter = ComObjectFactory.CreatePackageMatchFilter();
        }

        public void Dispose()
        {
        }

        public Guid Id => _guid;

        public string Name => "Windows Package Manager - WinGet";

        public string Description => "Finds missing commands that can be installed via WinGet.";

        /// <summary>
        /// Gets feedback based on the given commandline and error record.
        /// </summary>
        public FeedbackItem? GetFeedback(string commandLine, ErrorRecord lastError, CancellationToken token)
        {
            if (lastError.FullyQualifiedErrorId == "CommandNotFoundException")
            {
                var target = (string)lastError.TargetObject;
                var package = _FindPackage(target);
                if (package is null)
                {
                    return null;
                }
                _suggestion = "winget install --id " + package.Id;
                return new FeedbackItem(
                    Name,
                    new List<string> { _suggestion }
                );
            }
            return null;
        }

        private void _ApplyPackageMatchFilter(PackageMatchField field, PackageFieldMatchOption matchOption, string query)
        {
            // Configure filter
            _packageMatchFilter.Field = field;
            _packageMatchFilter.Option = matchOption;
            _packageMatchFilter.Value = query;

            // Apply filter
            _findPackagesOptions.Filters.Clear();
            _findPackagesOptions.Filters.Add(_packageMatchFilter);
        }

        private CatalogPackage? _TryGetBestMatchingPackage(IReadOnlyList<MatchResult> matches)
        {
            if (matches.Count == 1)
            {
                // One match --> return the package
                return matches.First().CatalogPackage;
            }
            else if (matches.Count > 1)
            {
                // Multiple matches --> return the one with the shortest match that starts with the query
                MatchResult? bestMatch = null;
                for (int i = 0; i < matches.Count; i++)
                {
                    var match = matches[i];
                    if (match.MatchCriteria.Option == PackageFieldMatchOption.EqualsCaseInsensitive || match.MatchCriteria.Option == PackageFieldMatchOption.Equals)
                    {
                        // Exact match --> return the package
                        return match.CatalogPackage;
                    }
                    else if (match.MatchCriteria.Option == PackageFieldMatchOption.StartsWithCaseInsensitive)
                    {
                        // get the shortest match that starts with the query
                        if (bestMatch == null || match.MatchCriteria.Value.Length < bestMatch.MatchCriteria.Value.Length)
                        {
                            bestMatch = match;
                        }
                    }
                }
                // bestMatch is null iff only ContainsCaseInsensitive matches exist.
                // There's no good way to figure out which one is relevant here, so just don't make a suggestion.
                return bestMatch == null ?
                    null :
                    bestMatch.CatalogPackage;
            }
            return null;
        }

        // Adapted from WinGet sample documentation: https://github.com/microsoft/winget-cli/blob/master/doc/specs/%23888%20-%20Com%20Api.md#32-search
        private CatalogPackage? _FindPackage(string query)
        {
            // Get the package catalog
            var catalogRef = _packageManager.GetPredefinedPackageCatalog(PredefinedPackageCatalog.OpenWindowsCatalog);
            var connectResult = catalogRef.Connect();
            byte retryCount = 0;
            while (connectResult.Status != ConnectResultStatus.Ok && retryCount < 3)
            {
                connectResult = catalogRef.Connect();
                retryCount++;
            }
            var catalog = connectResult.PackageCatalog;

            // Perform the query (search by command)
            _ApplyPackageMatchFilter(PackageMatchField.Command, PackageFieldMatchOption.StartsWithCaseInsensitive, query);
            var findPackagesResult = catalog.FindPackages(_findPackagesOptions);
            var matches = findPackagesResult.Matches;
            var pkg = _TryGetBestMatchingPackage(matches);
            if (pkg != null)
            {
                return pkg;
            }

            // No matches found when searching by command,
            // let's try again and search by name
            _ApplyPackageMatchFilter(PackageMatchField.Name, PackageFieldMatchOption.ContainsCaseInsensitive, query);

            // Perform the query (search by name)
            findPackagesResult = catalog.FindPackages(_findPackagesOptions);
            matches = findPackagesResult.Matches;
            return _TryGetBestMatchingPackage(matches);
        }

        public bool CanAcceptFeedback(PredictionClient client, PredictorFeedbackKind feedback)
        {
            return feedback switch
            {
                PredictorFeedbackKind.CommandLineAccepted => true,
                _ => false,
            };
        }

        public SuggestionPackage GetSuggestion(PredictionClient client, PredictionContext context, CancellationToken cancellationToken)
        {
            List<PredictiveSuggestion>? result = null;

            result ??= new List<PredictiveSuggestion>(1);
            if (_suggestion is null)
            {
                return default;
            }

            result.Add(new PredictiveSuggestion(_suggestion));

            if (result is not null)
            {
                return new SuggestionPackage(result);
            }

            return default;
        }

        public void OnCommandLineAccepted(PredictionClient client, IReadOnlyList<string> history)
        {
            _suggestion = null;
        }

        public void OnSuggestionDisplayed(PredictionClient client, uint session, int countOrIndex) { }

        public void OnSuggestionAccepted(PredictionClient client, uint session, string acceptedSuggestion) { }

        public void OnCommandLineExecuted(PredictionClient client, string commandLine, bool success) { }
    }
}