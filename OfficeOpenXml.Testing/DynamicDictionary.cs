using System.ComponentModel;
using System.Dynamic;
using System.Reflection;

namespace OfficeOpenXml.Testing;

/// <summary>
/// Class DynamicDictionary.
/// </summary>
/// <remarks>The class derived from DynamicObject. </remarks>
public class DynamicDictionary : DynamicObject
{
    private readonly Dictionary<string, object> _dictionary = new();

    /// <summary>
    /// Provides the implementation for operations that get member values. Classes derived from the <see cref="DynamicObject" /> class can override this method to specify dynamic behavior for operations such as getting a value for a property.
    /// </summary>
    /// <param name="binder">Provides information about the object that called the dynamic operation. The <c>binder.Name</c> property provides the name of the member on which the dynamic operation is performed. For example, for the <c>Console.WriteLine(sampleObject.SampleProperty)</c> statement, where <c>sampleObject</c> is an instance of the class derived from the <see cref="DynamicObject" /> class, <c>binder.Name</c> returns "SampleProperty". The <c>binder.IgnoreCase</c> property specifies whether the member name is case-sensitive.</param>
    /// <param name="result">The result of the get operation. For example, if the method is called for a property, you can assign the property value to <paramref name="result" />.</param>
    /// <returns><see langword="true" /> if the operation is successful; otherwise, <see langword="false" />. If this method returns <see langword="false" />, the run-time binder of the language determines the behavior. (In most cases, a run-time exception is thrown.)</returns>
    public override bool TryGetMember(GetMemberBinder binder, out object result)
    {
        string name = binder.Name;

        // If the property name is found in a dictionary,
        // set the result parameter to the property value and return true.
        // Otherwise, return false.
        return _dictionary.TryGetValue(name, out result);
    }

    /// <summary>
    /// Provides the implementation for operations that set member values. Classes derived from the <see cref="DynamicObject" /> class can override this method to specify dynamic behavior for operations such as setting a value for a property.
    /// </summary>
    /// <param name="binder">Provides information about the object that called the dynamic operation. The binder.Name property provides the name of the member to which the value is being assigned. For example, for the statement sampleObject.SampleProperty = "Test", where sampleObject is an instance of the class derived from the <see cref="DynamicObject" /> class, binder.Name returns "SampleProperty". The binder.IgnoreCase property specifies whether the member name is case-sensitive.</param>
    /// <param name="value">The value to set to the member. For example, for sampleObject.SampleProperty = "Test", where sampleObject is an instance of the class derived from the <see cref="DynamicObject" /> class, the <paramref name="value" /> is "Test".</param>
    /// <remarks>If you try to set a value of a property that is not defined in the class, this method is called. </remarks>
    /// <returns>true if the operation is successful; otherwise, false. If this method returns false, the run-time binder of the language determines the behavior. (In most cases, a language-specific run-time exception is thrown.)</returns>
    public override bool TrySetMember(SetMemberBinder binder, object value)
    {
        _dictionary[binder.Name] = value;

        // You can always add a value to a dictionary,
        // so this method always returns true.
        return true;
    }

    /// <summary>
    /// If a property value is a delegate, invoke it
    /// </summary>
    /// <param name="binder">Provides information about the dynamic operation. The binder.Name property provides the name of the member on which the dynamic operation is performed. For example, for the statement sampleObject.SampleMethod(100), where sampleObject is an instance of the class derived from the <see cref="DynamicObject" /> class, binder.Name returns "SampleMethod". The binder.IgnoreCase property specifies whether the member name is case-sensitive.</param>
    /// <param name="args">The arguments that are passed to the object member during the invoke operation. For example, for the statement sampleObject.SampleMethod(100), where sampleObject is derived from the <see cref="DynamicObject" /> class, <paramref name="args" /> is equal to 100.</param>
    /// <param name="result">The result of the member invocation.</param>
    /// <returns>true if the operation is successful; otherwise, false. If this method returns false, the run-time binder of the language determines the behavior. (In most cases, a language-specific run-time exception is thrown.)</returns>
    public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
    {
        if (_dictionary.TryGetValue(binder.Name, out object value) && value is Delegate del)
        {
            result = del.DynamicInvoke(args);
            return true;
        }

        return base.TryInvokeMember(binder, args, out result);
    }

    /// <summary>
    /// Return all dynamic member names
    /// </summary>
    /// <returns>A sequence that contains dynamic member names.</returns>
    public override IEnumerable<string> GetDynamicMemberNames()
    {
        return _dictionary.Keys;
    }

    /// <summary>
    /// Gets the dynamic values.
    /// </summary>
    /// <returns>ICollection&lt;System.Object&gt;.</returns>
    public ICollection<object> GetDynamicValues()
    {
        return _dictionary.Values;
    }

    /// <summary>
    /// Gets the member value.
    /// </summary>
    /// <param name="memberName">Name of the member.</param>
    /// <returns>System.Object.</returns>
    public object GetMemberValue(string memberName)
    {
        if (!_dictionary.TryGetValue(memberName, out object result))
        {
            Type valueType = GetType();
            PropertyInfo propertyInfo = valueType.GetProperty(memberName);

            if (propertyInfo != null)
            {
                result = propertyInfo.GetValue(this, null);
            }
        }

        return result;
    }

    /// <summary>
    /// Gets the member string value.
    /// </summary>
    /// <param name="memberName">Name of the member.</param>
    /// <returns>System.String.</returns>
    public string GetMemberStringValue(string memberName)
    {
        object result = GetMemberValue(memberName);

        if (result != null && result != DBNull.Value)
        {
            string value = result.ToString();

            if (value?.IsNullOrWhiteSpace() == false)
            {
                return value.Trim();
            }
        }

        return string.Empty;
    }

    /// <summary>
    /// Gets the properties members.
    /// </summary>
    /// <returns>IList&lt;System.String&gt;.</returns>
    public IList<string> GetPropertiesMembers()
    {
        PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(GetType());
        List<string> values = (from PropertyDescriptor property in properties select property.Name).ToList();

        foreach (string key in _dictionary.Keys)
        {
            if (!values.Contains(key))
            {
                values.Add(key);
            }
        }

        return values;
    }

    /// <summary>
    /// Sets the dynamic member.
    /// </summary>
    /// <param name="name">The name.</param>
    /// <param name="value">The value.</param>
    public void SetDynamicMember(string name, object value)
    {
        _dictionary[name] = value;
    }
}