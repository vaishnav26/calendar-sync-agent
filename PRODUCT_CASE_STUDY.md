# Product case study: Campus Calendar Sync

## User and job to be done

A student wants their own classes and timetable changes to appear in a calendar they already check. The source of truth is a shared timetable; the personal calendar is the consumption interface.

## Scope of the current solution

The workflow separates enrollment matching, timetable parsing, preview and explicit sync. The preview lets a student inspect the proposed classes before Calendar writes. Google sign-in determines the destination account; the enrollment workbook determines the course and section filters.

The public repository packages the existing Flask implementation for review. It replaces institutional data with synthetic data and provides an offline walkthrough. It does not change the deployed original.

## Trade-offs visible in the implementation

| Decision | Benefit | Cost or limitation |
|---|---|---|
| Explicit preview and sync | User sees the workflow and controls writes | Changes require another sync |
| Five-minute timetable cache | Lower quota use during bursts | Changes may take up to the cache window to appear |
| Deterministic event IDs | Repeated identical syncs avoid duplicate inserts | A new time produces a new identity; reschedules need reconciliation |
| Keep cancelled classes with a label | Retains context and removes reminders | Calendar remains visually populated |
| One instance with local SQLite | Simple deployment and attempt counting | Ephemeral state and limited horizontal scaling |
| Deterministic parsing | Predictable behavior with no model cost | Depends on the source timetable format |

## Validation included

Nine local automated tests cover cancellation formatting and core Calendar behavior. The offline demo shows the event payloads with synthetic data. Live OAuth, source permissions, timetable completeness, concurrency under real traffic and student acceptance still require separate validation.

## Metrics to measure next

These are proposed measurements, not achieved results:

- Activation: share of authorized users completing their first successful sync.
- Reliability: successful syncs divided by attempted syncs, grouped by failure reason.
- Correctness: sampled events matching source date, time, section and cancellation status.
- Duplicate rate: unintended duplicate events per audited sync.
- Freshness: delay from a source timetable change to the corresponding calendar update.
- User value: median time spent maintaining a timetable before and after using the workflow.

## Prioritized improvements

1. Reconcile rescheduled and removed classes while touching only app-managed events.
2. Make date filtering timezone-aware and test week/year transitions.
3. Verify enrollment identity and minimize Google API permissions and session exposure.
4. Move counters and session state to durable shared storage before scaling across instances.
5. Observe real usage and error rates before adding background sync or notifications.

## Where AI could help later

An AI-assisted parser could propose structured events from inconsistent timetable documents. That is a future idea, not an implemented feature. Evaluation would need a labelled set of timetable examples, field-level accuracy checks, cancellation recall, abstention behavior, cost and latency measurements, and human confirmation before Calendar writes. The current format already has a deterministic parser, so adding an LLM should first demonstrate a measurable benefit.
