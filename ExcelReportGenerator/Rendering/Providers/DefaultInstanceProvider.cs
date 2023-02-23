namespace ExcelReportGenerator.Rendering.Providers;

/// <summary>
/// Default implementation of <see cref="IInstanceProvider" />
/// </summary>
public class DefaultInstanceProvider : IInstanceProvider
{
    private readonly IDictionary<Type, object> _instanceCache = new Dictionary<Type, object>();

    /// <param name="defaultInstance">Instance which will be returned if the type is not specified explicitly</param>
    public DefaultInstanceProvider(object defaultInstance = null)
    {
        DefaultInstance = defaultInstance;
        if (DefaultInstance != null)
        {
            _instanceCache[DefaultInstance.GetType()] = DefaultInstance;
        }
    }

    /// <summary>
    /// Return instance if the type is not specified explicitly
    /// </summary>
    protected object DefaultInstance { get; }

    /// <inheritdoc />
    /// <summary>
    /// Provides instance of specified <paramref name="type"/> as singleton. Type must have a default constructor.
    /// </summary>
    /// <exception cref="InvalidOperationException"></exception>
    public virtual object GetInstance(Type type)
    {
        if (type == null)
        {
            return DefaultInstance ?? throw new InvalidOperationException("Type is not specified but defaultInstance is null");
        }

        if (_instanceCache.TryGetValue(type, out object instance))
        {
            return instance;
        }

        instance = Activator.CreateInstance(type);
        _instanceCache[type] = instance;
        return instance;
    }

    /// <inheritdoc />
    /// <summary>
    /// Provides instance of type <typeparamref name="T"/> as singleton. Type must have a default constructor.
    /// </summary>
    /// <exception cref="InvalidOperationException"></exception>
    public virtual T GetInstance<T>()
    {
        return (T)GetInstance(typeof(T));
    }
}