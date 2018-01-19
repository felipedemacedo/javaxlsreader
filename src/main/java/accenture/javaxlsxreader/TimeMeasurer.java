package accenture.javaxlsxreader;

public class TimeMeasurer {
	long startTime;
	long endTime;
	long elapsedTime = 0;
	
	TimeMeasurer(){
		start();
	}
	
	private long getTimeNow() {
		return System.nanoTime();
	}
	
	public void start() {
		setStartTime(getTimeNow());
	}
	
	public void stop() {
		setEndTime(getTimeNow());
		setElapsedTime(getEndTime() - getStartTime());
		System.out.println("ELAPSED TIME: " + getElapsedTimeInSeconds() + " s");  
	}
	
	public void restart() {
		stop();
		System.out.println("ELAPSED TIME: " + getElapsedTimeInSeconds() + " s");  
		start();
	}

	private long getStartTime() {
		return startTime;
	}

	private void setStartTime(long startTime) {
		this.startTime = startTime;
	}

	private long getEndTime() {
		return endTime;
	}

	private void setEndTime(long endTime) {
		this.endTime = endTime;
	}

	private long getElapsedTime() {
		return elapsedTime;
	}

	private void setElapsedTime(long elapsedTime) {
		this.elapsedTime = elapsedTime;
	}
	
	public double getElapsedTimeInSeconds() {
		double elapsedTimeInSecs = (double) getElapsedTime() / 1000000000.0;
		return elapsedTimeInSecs;
	}
}
