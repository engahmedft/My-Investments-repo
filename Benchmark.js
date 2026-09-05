// =========================================================================
// PERFORMANCE BENCHMARK UTILITY (Optimized)
// =========================================================================

/**
 * High-precision benchmark utility for tracking pipeline step performance.
 */
class Benchmark {
  /**
   * @param {string} pipelineName - Name of the pipeline being benchmarked.
   * @param {boolean} [enabled=true] - Enable switch for tracking and logging.
   * @param {number} [precision=2] - Number of decimal places for formatted durations.
   */
  constructor(pipelineName, enabled = true, precision = 2) {
    this.pipelineName = pipelineName;
    this.enabled = enabled;
    this.precision = precision;

    this.steps = [];
    this.startTime = this._now();

    if (this.enabled) {
      this._log(`\n🚀 === STARTING PIPELINE: ${this.pipelineName} ===`);
    }
  }

  /**
   * Monotonic high-resolution clock with Date.now fallback.
   * @private
   */
  _now() {
    return typeof performance !== 'undefined' ? performance.now() : Date.now();
  }

  /**
   * Formats duration in milliseconds to seconds string.
   * @private
   */
  _formatMs(ms) {
    return (ms / 1000).toFixed(this.precision) + 's';
  }

  /**
   * Console logger guarded by enabled flag.
   * @private
   */
  _log(...args) {
    if (this.enabled) console.log(...args);
  }

  /**
   * Console error logger guarded by enabled flag.
   * @private
   */
  _error(...args) {
    if (this.enabled) console.error(...args);
  }

  /** Enable benchmark tracking and logging */
  enable() {
    this.enabled = true;
    return this;
  }

  /** Disable benchmark tracking and logging */
  disable() {
    this.enabled = false;
    return this;
  }

  /**
   * Resets benchmark step history and start time.
   */
  reset() {
    this.steps = [];
    this.startTime = this._now();
    this._log(`\n🔄 === RESET PIPELINE: ${this.pipelineName} ===`);
    return this;
  }

  /**
   * Times the execution of a synchronous or asynchronous function.
   * When benchmark is disabled, executes `fn()` directly with near-zero overhead.
   * 
   * @param {string} label - Description of the step being benchmarked.
   * @param {Function} fn - Function to execute and measure.
   * @returns {Promise<*>|*} Result of `fn()`.
   */
  time(label, fn) {
    if (typeof fn !== 'function') {
      throw new TypeError(`Benchmark.time expected a function, got ${typeof fn}`);
    }

    // FAST PATH: If benchmark is disabled, bypass all timing & logging overhead completely!
    if (!this.enabled) {
      return fn();
    }

    this._log(`⏱️  [START] ${label}`);
    const start = this._now();

    const handleSuccess = (result) => {
      const elapsedMs = this._now() - start;
      this.steps.push({ label, durationMs: elapsedMs, failed: false });
      this._log(`✅ [DONE]  ${label} (${this._formatMs(elapsedMs)})`);
      return result;
    };

    const handleError = (error) => {
      const elapsedMs = this._now() - start;
      this.steps.push({ label, durationMs: elapsedMs, failed: true });
      this._error(`❌ [FAILED] ${label} after ${this._formatMs(elapsedMs)} - Error: ${error.message}`);
      throw error;
    };

    try {
      const result = fn();
      
      // Handle Promises / Async functions seamlessly
      if (result && typeof result.then === 'function') {
        return result.then(handleSuccess).catch(handleError);
      }
      
      return handleSuccess(result);
    } catch (error) {
      handleError(error);
    }
  }

  /**
   * Generates benchmark report and returns structured metric results.
   * @returns {{pipelineName: string, enabled: boolean, totalDurationMs: number, totalDurationSec: string, steps: Array}}
   */
  report() {
    if (!this.enabled) {
      return {
        pipelineName: this.pipelineName,
        enabled: false,
        totalDurationMs: 0,
        totalDurationSec: '0s',
        steps: []
      };
    }

    const totalMs = this._now() - this.startTime;
    const formattedTotal = this._formatMs(totalMs);

    const maxLabelLen = Math.max(38, ...this.steps.map(s => s.label.length));
    const divider = '='.repeat(maxLabelLen + 20);

    console.log(`\n${divider}`);
    console.log(`📊 PERFORMANCE BENCHMARK REPORT: ${this.pipelineName}`);
    console.log(`⏱️ Total Execution Time: ${formattedTotal}`);
    console.log('-'.repeat(maxLabelLen + 20));

    this.steps.forEach(step => {
      const paddedLabel = step.label.padEnd(maxLabelLen);
      const status = step.failed ? '❌ FAILED' : '  ';
      console.log(` - ${paddedLabel} : ${this._formatMs(step.durationMs)} ${status}`);
    });

    console.log(`${divider}\n`);

    return {
      pipelineName: this.pipelineName,
      enabled: true,
      totalDurationMs: totalMs,
      totalDurationSec: formattedTotal,
      steps: this.steps
    };
  }
}