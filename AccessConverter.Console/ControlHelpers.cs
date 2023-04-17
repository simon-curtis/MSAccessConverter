namespace AccessConverter.Console;

internal static class ControlHelpers
{
    public static T? GetProperty<T>(this Access.Control control, string name)
    {
        try
        {
            var property = (AccessDao.Property)control.Properties[name];
            return property.Value switch
            {
                T value => value,
                { } value => throw new Exception(
                    $"Type mismatch for property '{name}', expected: '{typeof(T)}' but got: '{value.GetType().Name}'")
            };
        }
        catch
        {
            return default;
        }
    }
}