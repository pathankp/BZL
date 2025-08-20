import "./logo.css"

export function Logo({ className }: { className?: string }) {
	return (
		<div className={className}>
			<h1 className="logo-text">ServerSentry</h1>
		</div>
	)
}
