namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}

namespace System.Runtime.CompilerServices
{
    [AttributeUsage(System.AttributeTargets.Class |
                           System.AttributeTargets.Struct |
                           System.AttributeTargets.Field |
                           System.AttributeTargets.Property,
                           AllowMultiple = false, Inherited = false)]
    internal sealed class RequiredMemberAttribute : Attribute
    {
        public RequiredMemberAttribute() { }
    }
}

namespace System.Diagnostics.CodeAnalysis
{
    [AttributeUsage(System.AttributeTargets.Constructor,
                           AllowMultiple = false, Inherited = false)]
    internal sealed class SetsRequiredMembersAttribute : Attribute
    {
        public SetsRequiredMembersAttribute() { }
    }
}

namespace System.Runtime.CompilerServices
{
    [AttributeUsage(
        System.AttributeTargets.Class |
        System.AttributeTargets.Struct |
        System.AttributeTargets.Method |
        System.AttributeTargets.Constructor |
        System.AttributeTargets.Property |
        System.AttributeTargets.Field,
        AllowMultiple = true,
        Inherited = false)]
    internal sealed class CompilerFeatureRequiredAttribute : Attribute
    {
        public CompilerFeatureRequiredAttribute(string featureName)
            => FeatureName = featureName;

        public string FeatureName { get; }
        public bool IsOptional { get; set; }
    }
}
