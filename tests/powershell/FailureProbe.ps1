# Explicit AC-009 fixture. Its name excludes it from *.Tests.ps1 discovery.
Describe 'Known failing Pester fixture' -Tag 'AC-009' {
    It 'must fail so the runner can prove nonzero propagation' {
        1 | Should-Be -Expected 2
    }
}
