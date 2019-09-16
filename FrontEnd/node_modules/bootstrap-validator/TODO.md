# TODO

###### BUGS
* Defer remote validation while request is still pending. (#72)

###### ENHANCEMENTS:
* Improve invalid/valid error events, add post-delay events.
  - [error/errored] [success/successed] in addition to [valid/invalid], upon displaying or clearing an error
  - add whether or not field is valid in [validated.bs.validator] event.detail
  - add events on `.validator('validate')`, including whole form validity in `event.detail`
  * ^ Add a way to reliably determine if form is valid or invalid upon submit. (#67)
* Refactor validators to optionally return promises. (#131) (#177) (#275)
* Add a class to the form to indicate validity state. (#260)


###### BREAKING CHANGES:
* Allow custom error messages to be functions
* Change remote validator to use response body as error message.


# DONE
