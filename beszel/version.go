package serversentry

import "github.com/blang/semver"

const (
	Version = "0.12.3"
	AppName = "ServerSentry"
)

var MinVersionCbor = semver.MustParse("0.12.0")
