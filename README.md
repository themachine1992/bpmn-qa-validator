# BPMN QA Validator

Automated quality-assurance checks for **BPMN 2.0** process models. Point it at an XML export (Signavio or any BPMN 2.0 tool) and get a prioritised list of modelling-standard violations, in the console or as a colour-coded Excel report.

Built out of real process-governance work: reviewing process models by hand against a modelling standard is slow and inconsistent, so the objective, rule-based part of that review was automated and the judgement-heavy part is left to the human reviewer.

## What it checks

29 rules (`R01`–`R29`), each with an ID, a severity and the standard criterion it maps to.

| Area | Rules |
|---|---|
| Swimlanes | empty lanes (R01), lane naming (R25), customer lane (R18), provider lanes (R20) |
| Tasks | task type set (R02), system attached (R03), verb-in-infinitive naming (R08), RPA status (R16), minimum activity count (R21) |
| Gateways | labelled outflows (R04), converging flows (R05), complex gateway (R22), parallel open/close balance (R27) |
| Events | start/end events and closing label (R06, R24), message flows through events (R09), escalation paths (R17) |
| Structure | disconnected nodes (R07), sub-processes (R10), unattached annotations (R28), blacklist of unsupported elements (R23) |
| Diagram metadata | mandatory attributes: regional code, business unit, process owner, description, criticality (R11–R15), frequency (R26), external portal (R19) |
| Metric | digitalisation / automation index (R29, informational) |

Every finding is tagged:

- `AUTO`: fully verifiable from the XML.
- `PARC`: partial evidence, a human must confirm.
- `INFO`: metric, not pass/fail.

## Usage

```bash
python bpmn_qa_validator.py examples/proceso_con_errores.bpmn
python bpmn_qa_validator.py examples/proceso_con_errores.bpmn --salida report.xlsx   # needs openpyxl
```

The console and Excel output is in Spanish (the language of the models it was built for); rule IDs and the code structure are language-neutral.

Sample output for the bundled example (15 findings):

```
[ALTA ][AUTO] R01  Lane (Actor) -> Calle vacia
              El actor/carril no tiene ninguna actividad asociada...
[ALTA ][AUTO] R07  Compuerta Compleja -> Complejo
              Elemento sin flujos de secuencia (desconectado del flujo).
[INFO ][INFO] R29  Proceso -> Solicitud de compra
              % Automatizacion = 0%  |  % Digitalizacion = 50% ...
```

## Adapting it to your own standard

The standard is data, not hard-coded logic. At the top of `bpmn_qa_validator.py`:

- `TAGS_PROHIBIDOS`: BPMN elements your standard does not allow.
- `ATRIBUTOS_OBLIGATORIOS_DIAGRAMA`: mandatory diagram attributes.
- `KEYS_ATRIBUTOS`: keywords used to detect those attributes in `extensionElements`.
- `NARANJA_HEX`: colour range used to detect the "external portal" convention.

Add a rule by creating a `Hallazgo(...)` inside `BpmnQAValidator.run()`.

## Tests

```bash
python -m unittest discover -s tests
```

## Design notes

- Standard library only (`xml.etree.ElementTree`); `openpyxl` is optional and used just for the Excel export.
- The validator indexes nodes, flows, lanes and BPMN-DI colours once, then each rule is a small independent pass over that index.
- Severity ordering is shared by the console and the report so reviewers always see blocking issues first.

## Author

Michel Urrutia Salgado · IT Project Manager & IT Governance · [LinkedIn](https://www.linkedin.com/in/michel-urrutia-salgado-9a4b3755/)

MIT licensed.
