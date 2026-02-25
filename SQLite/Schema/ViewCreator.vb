Module ViewCreator
  Friend Sub CreateViews(conn As SQLiteConnectionWrapper, manifest As SchemaManifest)

    Dim orderedViews = SchemaViewDependencyResolver.ResolveViewCreationOrder(manifest)

    For Each vw In orderedViews
      Dim sql As String = LoadEmbeddedTextResource(vw.SqlResource)
      conn.Execute($"DROP VIEW IF EXISTS {vw.Name};")
      conn.Execute(sql)
    Next

  End Sub

End Module
