
This is a back-port of GTranslate to .NET Framework 4.8.

The reason this was created is to test GTranslate in a legacy 4.8 project
that is also STA.

So the `AggregateTranslator` constructor has an optional bool parameter that
is passed into all calls to `ConfigureAwait`. If set to `true` then it is
assumed that execution is within an STA thread.
