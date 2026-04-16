use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{collect_range_values, evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "COUNT", description: "Counts cells containing numbers", syntax: "COUNT(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_count });
    reg.register(FnSpec { name: "COUNTA", description: "Counts non-empty cells", syntax: "COUNTA(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_counta });
    reg.register(FnSpec { name: "COUNTBLANK", description: "Counts empty cells", syntax: "COUNTBLANK(range)", min_args: 1, max_args: Some(1), eval: fn_countblank });
    reg.register(FnSpec { name: "COUNTIF", description: "Counts cells matching a condition", syntax: "COUNTIF(range, criteria)", min_args: 2, max_args: Some(2), eval: fn_countif });
    reg.register(FnSpec { name: "SUMIF", description: "Sums cells matching a condition", syntax: "SUMIF(range, criteria, [sum_range])", min_args: 2, max_args: Some(3), eval: fn_sumif });
    reg.register(FnSpec { name: "AVERAGEIF", description: "Averages cells matching a condition", syntax: "AVERAGEIF(range, criteria, [average_range])", min_args: 2, max_args: Some(3), eval: fn_averageif });
    reg.register(FnSpec { name: "MEDIAN", description: "Returns the median value", syntax: "MEDIAN(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_median });
    reg.register(FnSpec { name: "MODE", description: "Returns the most frequent value", syntax: "MODE(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_mode });
    reg.register(FnSpec { name: "STDEV", description: "Returns sample standard deviation", syntax: "STDEV(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_stdev });
    reg.register(FnSpec { name: "STDEVP", description: "Returns population standard deviation", syntax: "STDEVP(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_stdevp });
    reg.register(FnSpec { name: "VAR", description: "Returns sample variance", syntax: "VAR(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_var });
    reg.register(FnSpec { name: "VARP", description: "Returns population variance", syntax: "VARP(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_varp });
    reg.register(FnSpec { name: "LARGE", description: "Returns the k-th largest value", syntax: "LARGE(array, k)", min_args: 2, max_args: Some(2), eval: fn_large });
    reg.register(FnSpec { name: "SMALL", description: "Returns the k-th smallest value", syntax: "SMALL(array, k)", min_args: 2, max_args: Some(2), eval: fn_small });
    reg.register(FnSpec { name: "RANK", description: "Returns the rank of a number", syntax: "RANK(number, ref, [order])", min_args: 2, max_args: Some(3), eval: fn_rank });
    reg.register(FnSpec { name: "PERCENTILE", description: "Returns the k-th percentile", syntax: "PERCENTILE(array, k)", min_args: 2, max_args: Some(2), eval: fn_percentile });
    reg.register(FnSpec { name: "CORREL", description: "Returns correlation coefficient", syntax: "CORREL(array1, array2)", min_args: 2, max_args: Some(2), eval: fn_correl });
    reg.register(FnSpec { name: "MINIFS", description: "Returns minimum with conditions", syntax: "MINIFS(min_range, criteria_range1, criteria1, ...)", min_args: 3, max_args: None, eval: fn_minifs });
    reg.register(FnSpec { name: "MAXIFS", description: "Returns maximum with conditions", syntax: "MAXIFS(max_range, criteria_range1, criteria1, ...)", min_args: 3, max_args: None, eval: fn_maxifs });
    reg.register(FnSpec { name: "SUMIFS", description: "Sums cells with multiple conditions", syntax: "SUMIFS(sum_range, criteria_range1, criteria1, ...)", min_args: 3, max_args: None, eval: fn_sumifs });
    reg.register(FnSpec { name: "COUNTIFS", description: "Counts cells with multiple conditions", syntax: "COUNTIFS(criteria_range1, criteria1, ...)", min_args: 2, max_args: None, eval: fn_countifs });
    reg.register(FnSpec { name: "AVERAGEIFS", description: "Averages cells with multiple conditions", syntax: "AVERAGEIFS(avg_range, criteria_range1, criteria1, ...)", min_args: 3, max_args: None, eval: fn_averageifs });
    reg.register(FnSpec { name: "QUARTILE", description: "Returns the quartile of a data set", syntax: "QUARTILE(array, quart)", min_args: 2, max_args: Some(2), eval: fn_quartile });
    reg.register(FnSpec { name: "PERCENTILE.INC", description: "Returns inclusive percentile", syntax: "PERCENTILE.INC(array, k)", min_args: 2, max_args: Some(2), eval: fn_percentile });
    reg.register(FnSpec { name: "PERCENTILE.EXC", description: "Returns exclusive percentile", syntax: "PERCENTILE.EXC(array, k)", min_args: 2, max_args: Some(2), eval: fn_percentile_exc });
    reg.register(FnSpec { name: "FREQUENCY", description: "Returns frequency distribution", syntax: "FREQUENCY(data_array, bins_array)", min_args: 2, max_args: Some(2), eval: fn_frequency });
    reg.register(FnSpec { name: "SLOPE", description: "Returns slope of linear regression", syntax: "SLOPE(known_ys, known_xs)", min_args: 2, max_args: Some(2), eval: fn_slope });
    reg.register(FnSpec { name: "INTERCEPT", description: "Returns intercept of linear regression", syntax: "INTERCEPT(known_ys, known_xs)", min_args: 2, max_args: Some(2), eval: fn_intercept });
    reg.register(FnSpec { name: "RSQ", description: "Returns R-squared of linear regression", syntax: "RSQ(known_ys, known_xs)", min_args: 2, max_args: Some(2), eval: fn_rsq });
    reg.register(FnSpec { name: "FORECAST", description: "Predicts value using linear regression", syntax: "FORECAST(x, known_ys, known_xs)", min_args: 3, max_args: Some(3), eval: fn_forecast });
    reg.register(FnSpec { name: "STEYX", description: "Returns standard error of regression", syntax: "STEYX(known_ys, known_xs)", min_args: 2, max_args: Some(2), eval: fn_steyx });
    reg.register(FnSpec { name: "SKEW", description: "Returns skewness of distribution", syntax: "SKEW(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_skew });
    reg.register(FnSpec { name: "KURT", description: "Returns kurtosis of distribution", syntax: "KURT(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_kurt });
    reg.register(FnSpec { name: "COVARIANCE.P", description: "Returns population covariance", syntax: "COVARIANCE.P(array1, array2)", min_args: 2, max_args: Some(2), eval: fn_covariance_p });
    reg.register(FnSpec { name: "COVARIANCE.S", description: "Returns sample covariance", syntax: "COVARIANCE.S(array1, array2)", min_args: 2, max_args: Some(2), eval: fn_covariance_s });
    reg.register(FnSpec { name: "GEOMEAN", description: "Returns geometric mean", syntax: "GEOMEAN(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_geomean });
    reg.register(FnSpec { name: "HARMEAN", description: "Returns harmonic mean", syntax: "HARMEAN(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_harmean });
    reg.register(FnSpec { name: "TRIMMEAN", description: "Returns mean excluding outliers", syntax: "TRIMMEAN(array, percent)", min_args: 2, max_args: Some(2), eval: fn_trimmean });
    reg.register(FnSpec { name: "AVEDEV", description: "Returns average of absolute deviations", syntax: "AVEDEV(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_avedev });
    reg.register(FnSpec { name: "DEVSQ", description: "Returns sum of squared deviations", syntax: "DEVSQ(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_devsq });
    reg.register(FnSpec { name: "STANDARDIZE", description: "Returns a normalized value", syntax: "STANDARDIZE(x, mean, standard_dev)", min_args: 3, max_args: Some(3), eval: fn_standardize });
    reg.register(FnSpec { name: "FISHER", description: "Returns Fisher transformation", syntax: "FISHER(x)", min_args: 1, max_args: Some(1), eval: fn_fisher });
    reg.register(FnSpec { name: "FISHERINV", description: "Returns inverse Fisher transformation", syntax: "FISHERINV(y)", min_args: 1, max_args: Some(1), eval: fn_fisherinv });
    reg.register(FnSpec { name: "PROB", description: "Returns probability of a range", syntax: "PROB(x_range, prob_range, lower_limit, [upper_limit])", min_args: 3, max_args: Some(4), eval: fn_prob });
    reg.register(FnSpec { name: "PERMUT", description: "Returns number of permutations", syntax: "PERMUT(number, number_chosen)", min_args: 2, max_args: Some(2), eval: fn_permut });
    reg.register(FnSpec { name: "PERMUTATIONA", description: "Returns permutations with repetition", syntax: "PERMUTATIONA(number, number_chosen)", min_args: 2, max_args: Some(2), eval: fn_permutationa });
    reg.register(FnSpec { name: "NORM.DIST", description: "Returns normal distribution", syntax: "NORM.DIST(x, mean, standard_dev, cumulative)", min_args: 4, max_args: Some(4), eval: fn_norm_dist });
    reg.register(FnSpec { name: "NORM.INV", description: "Returns inverse normal distribution", syntax: "NORM.INV(probability, mean, standard_dev)", min_args: 3, max_args: Some(3), eval: fn_norm_inv });
    reg.register(FnSpec { name: "NORM.S.DIST", description: "Returns standard normal distribution", syntax: "NORM.S.DIST(z, cumulative)", min_args: 2, max_args: Some(2), eval: fn_norm_s_dist });
    reg.register(FnSpec { name: "NORM.S.INV", description: "Returns inverse standard normal", syntax: "NORM.S.INV(probability)", min_args: 1, max_args: Some(1), eval: fn_norm_s_inv });
    reg.register(FnSpec { name: "T.DIST", description: "Returns Student's t-distribution", syntax: "T.DIST(x, deg_freedom, cumulative)", min_args: 3, max_args: Some(3), eval: fn_t_dist });
    reg.register(FnSpec { name: "T.DIST.2T", description: "Returns two-tailed t-distribution", syntax: "T.DIST.2T(x, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_t_dist_2t });
    reg.register(FnSpec { name: "T.DIST.RT", description: "Returns right-tailed t-distribution", syntax: "T.DIST.RT(x, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_t_dist_rt });
    reg.register(FnSpec { name: "T.INV", description: "Returns inverse t-distribution", syntax: "T.INV(probability, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_t_inv });
    reg.register(FnSpec { name: "T.INV.2T", description: "Returns two-tailed inverse t-distribution", syntax: "T.INV.2T(probability, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_t_inv_2t });
    reg.register(FnSpec { name: "BINOM.DIST", description: "Returns binomial distribution", syntax: "BINOM.DIST(number_s, trials, probability_s, cumulative)", min_args: 4, max_args: Some(4), eval: fn_binom_dist });
    reg.register(FnSpec { name: "BINOM.INV", description: "Returns inverse binomial distribution", syntax: "BINOM.INV(trials, probability_s, alpha)", min_args: 3, max_args: Some(3), eval: fn_binom_inv });
    reg.register(FnSpec { name: "POISSON.DIST", description: "Returns Poisson distribution", syntax: "POISSON.DIST(x, mean, cumulative)", min_args: 3, max_args: Some(3), eval: fn_poisson_dist });
    reg.register(FnSpec { name: "EXPON.DIST", description: "Returns exponential distribution", syntax: "EXPON.DIST(x, lambda, cumulative)", min_args: 3, max_args: Some(3), eval: fn_expon_dist });
    reg.register(FnSpec { name: "GAMMA", description: "Returns the gamma function value", syntax: "GAMMA(number)", min_args: 1, max_args: Some(1), eval: fn_gamma });
    reg.register(FnSpec { name: "GAMMALN", description: "Returns ln of gamma function", syntax: "GAMMALN(x)", min_args: 1, max_args: Some(1), eval: fn_gammaln });
    reg.register(FnSpec { name: "GAMMA.DIST", description: "Returns gamma distribution", syntax: "GAMMA.DIST(x, alpha, beta, cumulative)", min_args: 4, max_args: Some(4), eval: fn_gamma_dist });
    reg.register(FnSpec { name: "GAMMA.INV", description: "Returns inverse gamma distribution", syntax: "GAMMA.INV(probability, alpha, beta)", min_args: 3, max_args: Some(3), eval: fn_gamma_inv });
    reg.register(FnSpec { name: "BETA.DIST", description: "Returns beta distribution", syntax: "BETA.DIST(x, alpha, beta, cumulative, [A], [B])", min_args: 4, max_args: Some(6), eval: fn_beta_dist });
    reg.register(FnSpec { name: "BETA.INV", description: "Returns inverse beta distribution", syntax: "BETA.INV(probability, alpha, beta, [A], [B])", min_args: 3, max_args: Some(5), eval: fn_beta_inv });
    reg.register(FnSpec { name: "WEIBULL.DIST", description: "Returns Weibull distribution", syntax: "WEIBULL.DIST(x, alpha, beta, cumulative)", min_args: 4, max_args: Some(4), eval: fn_weibull_dist });
    reg.register(FnSpec { name: "LOGNORM.DIST", description: "Returns lognormal distribution", syntax: "LOGNORM.DIST(x, mean, standard_dev, cumulative)", min_args: 4, max_args: Some(4), eval: fn_lognorm_dist });
    reg.register(FnSpec { name: "LOGNORM.INV", description: "Returns inverse lognormal distribution", syntax: "LOGNORM.INV(probability, mean, standard_dev)", min_args: 3, max_args: Some(3), eval: fn_lognorm_inv });
    reg.register(FnSpec { name: "CHISQ.DIST", description: "Returns chi-squared distribution", syntax: "CHISQ.DIST(x, deg_freedom, cumulative)", min_args: 3, max_args: Some(3), eval: fn_chisq_dist });
    reg.register(FnSpec { name: "CHISQ.DIST.RT", description: "Returns right-tailed chi-squared", syntax: "CHISQ.DIST.RT(x, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_chisq_dist_rt });
    reg.register(FnSpec { name: "CHISQ.INV", description: "Returns inverse chi-squared", syntax: "CHISQ.INV(probability, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_chisq_inv });
    reg.register(FnSpec { name: "CHISQ.INV.RT", description: "Returns right-tailed inverse chi-squared", syntax: "CHISQ.INV.RT(probability, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_chisq_inv_rt });
    reg.register(FnSpec { name: "F.DIST", description: "Returns F probability distribution", syntax: "F.DIST(x, deg_freedom1, deg_freedom2, cumulative)", min_args: 4, max_args: Some(4), eval: fn_f_dist });
    reg.register(FnSpec { name: "F.DIST.RT", description: "Returns right-tailed F distribution", syntax: "F.DIST.RT(x, deg_freedom1, deg_freedom2)", min_args: 3, max_args: Some(3), eval: fn_f_dist_rt });
    reg.register(FnSpec { name: "F.INV", description: "Returns inverse F distribution", syntax: "F.INV(probability, deg_freedom1, deg_freedom2)", min_args: 3, max_args: Some(3), eval: fn_f_inv });
    reg.register(FnSpec { name: "F.INV.RT", description: "Returns right-tailed inverse F", syntax: "F.INV.RT(probability, deg_freedom1, deg_freedom2)", min_args: 3, max_args: Some(3), eval: fn_f_inv_rt });
    reg.register(FnSpec { name: "CONFIDENCE.NORM", description: "Returns confidence interval using normal", syntax: "CONFIDENCE.NORM(alpha, standard_dev, size)", min_args: 3, max_args: Some(3), eval: fn_confidence_norm });
    reg.register(FnSpec { name: "CONFIDENCE.T", description: "Returns confidence interval using t", syntax: "CONFIDENCE.T(alpha, standard_dev, size)", min_args: 3, max_args: Some(3), eval: fn_confidence_t });
    reg.register(FnSpec { name: "STDEV.S", description: "Returns sample standard deviation", syntax: "STDEV.S(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_stdev });
    reg.register(FnSpec { name: "STDEV.P", description: "Returns population standard deviation", syntax: "STDEV.P(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_stdevp });
    reg.register(FnSpec { name: "VAR.S", description: "Returns sample variance", syntax: "VAR.S(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_var });
    reg.register(FnSpec { name: "VAR.P", description: "Returns population variance", syntax: "VAR.P(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_varp });
    reg.register(FnSpec { name: "MODE.SNGL", description: "Returns most common value", syntax: "MODE.SNGL(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_mode });
    reg.register(FnSpec { name: "RANK.AVG", description: "Returns rank with average for ties", syntax: "RANK.AVG(number, ref, [order])", min_args: 2, max_args: Some(3), eval: fn_rank_avg });
    reg.register(FnSpec { name: "RANK.EQ", description: "Returns rank of a number", syntax: "RANK.EQ(number, ref, [order])", min_args: 2, max_args: Some(3), eval: fn_rank });
    reg.register(FnSpec { name: "QUARTILE.INC", description: "Returns inclusive quartile", syntax: "QUARTILE.INC(array, quart)", min_args: 2, max_args: Some(2), eval: fn_quartile });
    reg.register(FnSpec { name: "QUARTILE.EXC", description: "Returns exclusive quartile", syntax: "QUARTILE.EXC(array, quart)", min_args: 2, max_args: Some(2), eval: fn_quartile_exc });
    reg.register(FnSpec { name: "SKEW.P", description: "Returns population skewness", syntax: "SKEW.P(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_skew_p });
    reg.register(FnSpec { name: "GROWTH", description: "Returns exponential growth values", syntax: "GROWTH(known_ys, [known_xs], [new_xs], [const])", min_args: 1, max_args: Some(4), eval: fn_growth });
    reg.register(FnSpec { name: "TREND", description: "Returns linear trend values", syntax: "TREND(known_ys, [known_xs], [new_xs], [const])", min_args: 1, max_args: Some(4), eval: fn_trend });
    reg.register(FnSpec { name: "FORECAST.LINEAR", description: "Predicts value using linear regression", syntax: "FORECAST.LINEAR(x, known_ys, known_xs)", min_args: 3, max_args: Some(3), eval: fn_forecast });
    reg.register(FnSpec { name: "PEARSON", description: "Returns Pearson correlation coefficient", syntax: "PEARSON(array1, array2)", min_args: 2, max_args: Some(2), eval: fn_correl });
    reg.register(FnSpec { name: "COMBIN", description: "Returns combinations", syntax: "COMBIN(number, number_chosen)", min_args: 2, max_args: Some(2), eval: fn_combin_stat });
    reg.register(FnSpec { name: "COMBINA", description: "Returns combinations with repetition", syntax: "COMBINA(number, number_chosen)", min_args: 2, max_args: Some(2), eval: fn_combina_stat });
    reg.register(FnSpec { name: "HYPGEOM.DIST", description: "Returns hypergeometric distribution", syntax: "HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)", min_args: 5, max_args: Some(5), eval: fn_hypgeom_dist });
    reg.register(FnSpec { name: "Z.TEST", description: "Returns two-tailed P-value of z-test", syntax: "Z.TEST(array, x, [sigma])", min_args: 2, max_args: Some(3), eval: fn_z_test });
    // Compatibility aliases (old Excel names pointing to new implementations)
    reg.register(FnSpec { name: "BETADIST", description: "Returns beta distribution (compat)", syntax: "BETADIST(x, alpha, beta, [A], [B])", min_args: 3, max_args: Some(5), eval: fn_betadist_compat });
    reg.register(FnSpec { name: "BETAINV", description: "Returns inverse beta (compat)", syntax: "BETAINV(probability, alpha, beta, [A], [B])", min_args: 3, max_args: Some(5), eval: fn_beta_inv });
    reg.register(FnSpec { name: "BINOMDIST", description: "Returns binomial distribution (compat)", syntax: "BINOMDIST(number_s, trials, probability_s, cumulative)", min_args: 4, max_args: Some(4), eval: fn_binom_dist });
    reg.register(FnSpec { name: "BINOM.DIST.RANGE", description: "Returns binomial distribution probability", syntax: "BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2])", min_args: 3, max_args: Some(4), eval: fn_binom_dist_range });
    reg.register(FnSpec { name: "CHIDIST", description: "Returns chi-squared right-tail (compat)", syntax: "CHIDIST(x, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_chisq_dist_rt });
    reg.register(FnSpec { name: "CHIINV", description: "Returns inverse chi-squared right-tail (compat)", syntax: "CHIINV(probability, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_chisq_inv_rt });
    reg.register(FnSpec { name: "CHISQ.TEST", description: "Returns chi-squared test statistic", syntax: "CHISQ.TEST(actual_range, expected_range)", min_args: 2, max_args: Some(2), eval: fn_chisq_test });
    reg.register(FnSpec { name: "CHITEST", description: "Returns chi-squared test (compat)", syntax: "CHITEST(actual_range, expected_range)", min_args: 2, max_args: Some(2), eval: fn_chisq_test });
    reg.register(FnSpec { name: "CONFIDENCE", description: "Returns confidence interval (compat)", syntax: "CONFIDENCE(alpha, standard_dev, size)", min_args: 3, max_args: Some(3), eval: fn_confidence_norm });
    reg.register(FnSpec { name: "COVAR", description: "Returns population covariance (compat)", syntax: "COVAR(array1, array2)", min_args: 2, max_args: Some(2), eval: fn_covariance_p });
    reg.register(FnSpec { name: "CRITBINOM", description: "Returns inverse binomial (compat)", syntax: "CRITBINOM(trials, probability_s, alpha)", min_args: 3, max_args: Some(3), eval: fn_binom_inv });
    reg.register(FnSpec { name: "EXPONDIST", description: "Returns exponential distribution (compat)", syntax: "EXPONDIST(x, lambda, cumulative)", min_args: 3, max_args: Some(3), eval: fn_expon_dist });
    reg.register(FnSpec { name: "F.TEST", description: "Returns F-test result", syntax: "F.TEST(array1, array2)", min_args: 2, max_args: Some(2), eval: fn_f_test });
    reg.register(FnSpec { name: "GAMMADIST", description: "Returns gamma distribution (compat)", syntax: "GAMMADIST(x, alpha, beta, cumulative)", min_args: 4, max_args: Some(4), eval: fn_gamma_dist });
    reg.register(FnSpec { name: "GAMMAINV", description: "Returns inverse gamma (compat)", syntax: "GAMMAINV(probability, alpha, beta)", min_args: 3, max_args: Some(3), eval: fn_gamma_inv });
    reg.register(FnSpec { name: "GAMMALN.PRECISE", description: "Returns ln of gamma function", syntax: "GAMMALN.PRECISE(x)", min_args: 1, max_args: Some(1), eval: fn_gammaln });
    reg.register(FnSpec { name: "GAUSS", description: "Returns 0.5 less than standard normal CDF", syntax: "GAUSS(z)", min_args: 1, max_args: Some(1), eval: fn_gauss });
    reg.register(FnSpec { name: "HYPGEOMDIST", description: "Returns hypergeometric distribution (compat)", syntax: "HYPGEOMDIST(sample_s, number_sample, population_s, number_pop)", min_args: 4, max_args: Some(4), eval: fn_hypgeomdist_compat });
    reg.register(FnSpec { name: "LINEST", description: "Returns linear regression statistics", syntax: "LINEST(known_ys, [known_xs], [const], [stats])", min_args: 1, max_args: Some(4), eval: fn_linest });
    reg.register(FnSpec { name: "LOGEST", description: "Returns exponential regression statistics", syntax: "LOGEST(known_ys, [known_xs], [const], [stats])", min_args: 1, max_args: Some(4), eval: fn_logest });
    reg.register(FnSpec { name: "LOGINV", description: "Returns inverse lognormal (compat)", syntax: "LOGINV(probability, mean, standard_dev)", min_args: 3, max_args: Some(3), eval: fn_lognorm_inv });
    reg.register(FnSpec { name: "LOGNORMDIST", description: "Returns lognormal distribution (compat)", syntax: "LOGNORMDIST(x, mean, standard_dev)", min_args: 3, max_args: Some(3), eval: fn_lognormdist_compat });
    reg.register(FnSpec { name: "AVERAGEA", description: "Averages values including text and logical", syntax: "AVERAGEA(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_averagea });
    reg.register(FnSpec { name: "MAXA", description: "Returns max including text and logical", syntax: "MAXA(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_maxa });
    reg.register(FnSpec { name: "MINA", description: "Returns min including text and logical", syntax: "MINA(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_mina });
    reg.register(FnSpec { name: "MODE.MULT", description: "Returns array of most frequent values", syntax: "MODE.MULT(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_mode_mult });
    reg.register(FnSpec { name: "NEGBINOM.DIST", description: "Returns negative binomial distribution", syntax: "NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)", min_args: 4, max_args: Some(4), eval: fn_negbinom_dist });
    reg.register(FnSpec { name: "NEGBINOMDIST", description: "Returns negative binomial (compat)", syntax: "NEGBINOMDIST(number_f, number_s, probability_s)", min_args: 3, max_args: Some(3), eval: fn_negbinomdist_compat });
    reg.register(FnSpec { name: "NORMDIST", description: "Returns normal distribution (compat)", syntax: "NORMDIST(x, mean, standard_dev, cumulative)", min_args: 4, max_args: Some(4), eval: fn_norm_dist });
    reg.register(FnSpec { name: "NORMINV", description: "Returns inverse normal (compat)", syntax: "NORMINV(probability, mean, standard_dev)", min_args: 3, max_args: Some(3), eval: fn_norm_inv });
    reg.register(FnSpec { name: "NORMSDIST", description: "Returns standard normal CDF (compat)", syntax: "NORMSDIST(z)", min_args: 1, max_args: Some(1), eval: fn_normsdist_compat });
    reg.register(FnSpec { name: "NORMSINV", description: "Returns inverse standard normal (compat)", syntax: "NORMSINV(probability)", min_args: 1, max_args: Some(1), eval: fn_norm_s_inv });
    reg.register(FnSpec { name: "PERCENTRANK", description: "Returns percentile rank", syntax: "PERCENTRANK(array, x, [significance])", min_args: 2, max_args: Some(3), eval: fn_percentrank });
    reg.register(FnSpec { name: "PERCENTRANK.INC", description: "Returns inclusive percentile rank", syntax: "PERCENTRANK.INC(array, x, [significance])", min_args: 2, max_args: Some(3), eval: fn_percentrank });
    reg.register(FnSpec { name: "PERCENTRANK.EXC", description: "Returns exclusive percentile rank", syntax: "PERCENTRANK.EXC(array, x, [significance])", min_args: 2, max_args: Some(3), eval: fn_percentrank_exc });
    reg.register(FnSpec { name: "PHI", description: "Returns standard normal PDF value", syntax: "PHI(x)", min_args: 1, max_args: Some(1), eval: fn_phi });
    reg.register(FnSpec { name: "POISSON", description: "Returns Poisson distribution (compat)", syntax: "POISSON(x, mean, cumulative)", min_args: 3, max_args: Some(3), eval: fn_poisson_dist });
    reg.register(FnSpec { name: "STDEVA", description: "Estimates stdev including text and logical", syntax: "STDEVA(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_stdeva });
    reg.register(FnSpec { name: "STDEVPA", description: "Calculates population stdev including text and logical", syntax: "STDEVPA(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_stdevpa });
    reg.register(FnSpec { name: "TDIST", description: "Returns t-distribution (compat)", syntax: "TDIST(x, deg_freedom, tails)", min_args: 3, max_args: Some(3), eval: fn_tdist_compat });
    reg.register(FnSpec { name: "TINV", description: "Returns inverse t-distribution (compat)", syntax: "TINV(probability, deg_freedom)", min_args: 2, max_args: Some(2), eval: fn_t_inv_2t });
    reg.register(FnSpec { name: "T.TEST", description: "Returns t-test probability", syntax: "T.TEST(array1, array2, tails, type)", min_args: 4, max_args: Some(4), eval: fn_t_test });
    reg.register(FnSpec { name: "TTEST", description: "Returns t-test probability (compat)", syntax: "TTEST(array1, array2, tails, type)", min_args: 4, max_args: Some(4), eval: fn_t_test });
    reg.register(FnSpec { name: "VARA", description: "Estimates variance including text and logical", syntax: "VARA(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_vara });
    reg.register(FnSpec { name: "VARPA", description: "Calculates population variance including text and logical", syntax: "VARPA(value1, [value2], ...)", min_args: 1, max_args: None, eval: fn_varpa });
    reg.register(FnSpec { name: "WEIBULL", description: "Returns Weibull distribution (compat)", syntax: "WEIBULL(x, alpha, beta, cumulative)", min_args: 4, max_args: Some(4), eval: fn_weibull_dist });
    reg.register(FnSpec { name: "ZTEST", description: "Returns z-test (compat)", syntax: "ZTEST(array, x, [sigma])", min_args: 2, max_args: Some(3), eval: fn_z_test });
    reg.register(FnSpec { name: "FORECAST.ETS", description: "Returns forecast using ETS algorithm", syntax: "FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])", min_args: 3, max_args: Some(6), eval: fn_forecast_ets_stub });
    reg.register(FnSpec { name: "FORECAST.ETS.CONFINT", description: "Returns ETS confidence interval", syntax: "FORECAST.ETS.CONFINT(target_date, values, timeline, [confidence_level], [seasonality])", min_args: 3, max_args: Some(5), eval: fn_forecast_ets_stub });
    reg.register(FnSpec { name: "FORECAST.ETS.SEASONALITY", description: "Returns ETS seasonality length", syntax: "FORECAST.ETS.SEASONALITY(values, timeline, [data_completion], [aggregation])", min_args: 2, max_args: Some(4), eval: fn_forecast_ets_stub });
    reg.register(FnSpec { name: "FORECAST.ETS.STAT", description: "Returns ETS statistical value", syntax: "FORECAST.ETS.STAT(values, timeline, statistic_type, [seasonality])", min_args: 3, max_args: Some(4), eval: fn_forecast_ets_stub });
}

fn collect_numbers(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Vec<f64> {
    let mut nums = Vec::new();
    for arg in args {
        match arg {
            Expr::Range { start, end } => {
                for val in collect_range_values(start, end, ctx) {
                    if let Some(n) = val.as_f64() {
                        nums.push(n);
                    }
                }
            }
            _ => {
                if let Some(n) = evaluate(arg, ctx, reg).as_f64() {
                    nums.push(n);
                }
            }
        }
    }
    nums
}

fn collect_all_values(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Vec<CellValue> {
    let mut vals = Vec::new();
    for arg in args {
        match arg {
            Expr::Range { start, end } => {
                vals.extend(collect_range_values(start, end, ctx));
            }
            _ => {
                vals.push(evaluate(arg, ctx, reg));
            }
        }
    }
    vals
}

fn fn_count(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let vals = collect_all_values(args, ctx, reg);
    let count = vals.iter().filter(|v| v.as_f64().is_some()).count();
    CellValue::Number(count as f64)
}

fn fn_counta(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let vals = collect_all_values(args, ctx, reg);
    let count = vals.iter().filter(|v| !matches!(v, CellValue::Empty)).count();
    CellValue::Number(count as f64)
}

fn fn_countblank(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let count = vals.iter().filter(|v| matches!(v, CellValue::Empty)).count();
    CellValue::Number(count as f64)
}

fn matches_criteria(val: &CellValue, criteria: &str) -> bool {
    if let Some(rest) = criteria.strip_prefix(">=") {
        if let (Some(vn), Ok(cn)) = (val.as_f64(), rest.parse::<f64>()) {
            return vn >= cn;
        }
    } else if let Some(rest) = criteria.strip_prefix("<=") {
        if let (Some(vn), Ok(cn)) = (val.as_f64(), rest.parse::<f64>()) {
            return vn <= cn;
        }
    } else if let Some(rest) = criteria.strip_prefix("<>") {
        return val.display_value() != rest;
    } else if let Some(rest) = criteria.strip_prefix('>') {
        if let (Some(vn), Ok(cn)) = (val.as_f64(), rest.parse::<f64>()) {
            return vn > cn;
        }
    } else if let Some(rest) = criteria.strip_prefix('<') {
        if let (Some(vn), Ok(cn)) = (val.as_f64(), rest.parse::<f64>()) {
            return vn < cn;
        }
    } else if let Some(rest) = criteria.strip_prefix('=') {
        return val.display_value().eq_ignore_ascii_case(rest);
    }

    if let Ok(cn) = criteria.parse::<f64>() {
        if let Some(vn) = val.as_f64() {
            return (vn - cn).abs() < f64::EPSILON;
        }
    }
    val.display_value().eq_ignore_ascii_case(criteria)
}

fn fn_countif(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let criteria = evaluate(&args[1], ctx, reg).display_value();
    let count = vals.iter().filter(|v| matches_criteria(v, &criteria)).count();
    CellValue::Number(count as f64)
}

fn fn_sumif(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let range_vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let criteria = evaluate(&args[1], ctx, reg).display_value();
    let sum_vals = if args.len() > 2 {
        match &args[2] {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => return CellValue::Error(CellError::Value),
        }
    } else {
        range_vals.clone()
    };

    let mut sum = 0.0;
    for (i, v) in range_vals.iter().enumerate() {
        if matches_criteria(v, &criteria) {
            if let Some(n) = sum_vals.get(i).and_then(|v| v.as_f64()) {
                sum += n;
            }
        }
    }
    CellValue::Number(sum)
}

fn fn_averageif(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let range_vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let criteria = evaluate(&args[1], ctx, reg).display_value();
    let sum_vals = if args.len() > 2 {
        match &args[2] {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => return CellValue::Error(CellError::Value),
        }
    } else {
        range_vals.clone()
    };

    let mut sum = 0.0;
    let mut count = 0;
    for (i, v) in range_vals.iter().enumerate() {
        if matches_criteria(v, &criteria) {
            if let Some(n) = sum_vals.get(i).and_then(|v| v.as_f64()) {
                sum += n;
                count += 1;
            }
        }
    }
    if count == 0 {
        CellValue::Error(CellError::Div0)
    } else {
        CellValue::Number(sum / count as f64)
    }
}

fn fn_median(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let mid = nums.len() / 2;
    if nums.len() % 2 == 0 {
        CellValue::Number((nums[mid - 1] + nums[mid]) / 2.0)
    } else {
        CellValue::Number(nums[mid])
    }
}

fn fn_mode(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Na);
    }
    let mut counts = std::collections::HashMap::new();
    for n in &nums {
        let key = n.to_bits();
        *counts.entry(key).or_insert(0) += 1;
    }
    let max_count = *counts.values().max().unwrap();
    if max_count == 1 {
        return CellValue::Error(CellError::Na);
    }
    for n in &nums {
        if counts[&n.to_bits()] == max_count {
            return CellValue::Number(*n);
        }
    }
    CellValue::Error(CellError::Na)
}

fn variance(nums: &[f64], sample: bool) -> Option<f64> {
    let n = nums.len();
    if n == 0 || (sample && n == 1) {
        return None;
    }
    let mean = nums.iter().sum::<f64>() / n as f64;
    let sum_sq: f64 = nums.iter().map(|x| (x - mean).powi(2)).sum();
    let divisor = if sample { (n - 1) as f64 } else { n as f64 };
    Some(sum_sq / divisor)
}

fn fn_stdev(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    match variance(&nums, true) {
        Some(v) => CellValue::Number(v.sqrt()),
        None => CellValue::Error(CellError::Div0),
    }
}

fn fn_stdevp(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    match variance(&nums, false) {
        Some(v) => CellValue::Number(v.sqrt()),
        None => CellValue::Error(CellError::Div0),
    }
}

fn fn_var(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    match variance(&nums, true) {
        Some(v) => CellValue::Number(v),
        None => CellValue::Error(CellError::Div0),
    }
}

fn fn_varp(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    match variance(&nums, false) {
        Some(v) => CellValue::Number(v),
        None => CellValue::Error(CellError::Div0),
    }
}

fn fn_large(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let k = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => n as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    if k > nums.len() {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| b.partial_cmp(a).unwrap());
    CellValue::Number(nums[k - 1])
}

fn fn_small(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let k = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => n as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    if k > nums.len() {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    CellValue::Number(nums[k - 1])
}

fn fn_rank(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let number = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) => n,
        None => return CellValue::Error(CellError::Value),
    };
    let nums = match &args[1] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let descending = if args.len() > 2 {
        evaluate(&args[2], ctx, reg).as_f64().unwrap_or(0.0) == 0.0
    } else {
        true
    };

    let rank = if descending {
        nums.iter().filter(|&&n| n > number).count() + 1
    } else {
        nums.iter().filter(|&&n| n < number).count() + 1
    };
    CellValue::Number(rank as f64)
}

fn fn_percentile(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let k = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if (0.0..=1.0).contains(&n) => n,
        _ => return CellValue::Error(CellError::Num),
    };
    if nums.is_empty() {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let n = nums.len() - 1;
    let idx = k * n as f64;
    let lower = idx.floor() as usize;
    let upper = idx.ceil() as usize;
    let frac = idx - lower as f64;
    let result = nums[lower] * (1.0 - frac) + nums[upper] * frac;
    CellValue::Number(result)
}

fn fn_correl(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let xs = match &args[0] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let ys = match &args[1] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let n = xs.len().min(ys.len());
    if n < 2 {
        return CellValue::Error(CellError::Na);
    }
    let mean_x = xs[..n].iter().sum::<f64>() / n as f64;
    let mean_y = ys[..n].iter().sum::<f64>() / n as f64;
    let mut sum_xy = 0.0;
    let mut sum_x2 = 0.0;
    let mut sum_y2 = 0.0;
    for i in 0..n {
        let dx = xs[i] - mean_x;
        let dy = ys[i] - mean_y;
        sum_xy += dx * dy;
        sum_x2 += dx * dx;
        sum_y2 += dy * dy;
    }
    let denom = (sum_x2 * sum_y2).sqrt();
    if denom == 0.0 {
        CellValue::Error(CellError::Div0)
    } else {
        CellValue::Number(sum_xy / denom)
    }
}

fn fn_minifs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    conditional_aggregate(args, ctx, reg, |vals| {
        vals.iter().cloned().fold(f64::INFINITY, f64::min)
    })
}

fn fn_maxifs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    conditional_aggregate(args, ctx, reg, |vals| {
        vals.iter().cloned().fold(f64::NEG_INFINITY, f64::max)
    })
}

fn conditional_aggregate(
    args: &[Expr],
    ctx: &dyn EvalContext,
    reg: &FunctionRegistry,
    agg: fn(&[f64]) -> f64,
) -> CellValue {
    let target_vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };

    let mut mask = vec![true; target_vals.len()];

    let mut i = 1;
    while i + 1 < args.len() {
        let criteria_range = match &args[i] {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => return CellValue::Error(CellError::Value),
        };
        let criteria = evaluate(&args[i + 1], ctx, reg).display_value();
        for (j, v) in criteria_range.iter().enumerate() {
            if j < mask.len() && !matches_criteria(v, &criteria) {
                mask[j] = false;
            }
        }
        i += 2;
    }

    let filtered: Vec<f64> = target_vals
        .iter()
        .enumerate()
        .filter(|(j, _)| mask.get(*j).copied().unwrap_or(false))
        .filter_map(|(_, v)| v.as_f64())
        .collect();

    if filtered.is_empty() {
        CellValue::Number(0.0)
    } else {
        CellValue::Number(agg(&filtered))
    }
}

fn fn_sumifs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    conditional_aggregate(args, ctx, reg, |vals| vals.iter().sum())
}

fn fn_countifs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let first_range = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let mut mask = vec![true; first_range.len()];
    let mut i = 0;
    while i + 1 < args.len() {
        let range = match &args[i] {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => return CellValue::Error(CellError::Value),
        };
        let criteria = evaluate(&args[i + 1], ctx, reg).display_value();
        for (j, v) in range.iter().enumerate() {
            if j < mask.len() && !matches_criteria(v, &criteria) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let count = mask.iter().filter(|&&m| m).count();
    CellValue::Number(count as f64)
}

fn fn_averageifs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    conditional_aggregate(args, ctx, reg, |vals| {
        if vals.is_empty() { 0.0 } else { vals.iter().sum::<f64>() / vals.len() as f64 }
    })
}

fn fn_quartile(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    let q = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if (0.0..=4.0).contains(&n) => n,
        _ => return CellValue::Error(CellError::Num),
    };
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let k = q / 4.0;
    let n = nums.len() - 1;
    let idx = k * n as f64;
    let lower = idx.floor() as usize;
    let upper = idx.ceil() as usize;
    let frac = idx - lower as f64;
    CellValue::Number(nums[lower] * (1.0 - frac) + nums[upper] * frac)
}

fn fn_percentile_exc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    let k = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) => n,
        None => return CellValue::Error(CellError::Num),
    };
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    let n = nums.len();
    if k <= 1.0 / (n as f64 + 1.0) || k >= n as f64 / (n as f64 + 1.0) {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let rank = k * (n as f64 + 1.0) - 1.0;
    let lower = rank.floor() as usize;
    let upper = (lower + 1).min(n - 1);
    let frac = rank - lower as f64;
    CellValue::Number(nums[lower] * (1.0 - frac) + nums[upper] * frac)
}

fn fn_frequency(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let data = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    let mut bins = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    bins.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let mut freq = vec![0.0; bins.len() + 1];
    for &d in &data {
        let mut placed = false;
        for (i, &b) in bins.iter().enumerate() {
            if d <= b { freq[i] += 1.0; placed = true; break; }
        }
        if !placed { freq[bins.len()] += 1.0; }
    }
    let rows: Vec<Vec<CellValue>> = freq.iter().map(|&f| vec![CellValue::Number(f)]).collect();
    CellValue::Array(Box::new(rows))
}

fn linear_regression(args: &[Expr], ctx: &dyn EvalContext) -> Option<(f64, f64, Vec<f64>, Vec<f64>)> {
    let ys = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return None,
    };
    let xs = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return None,
    };
    let n = xs.len().min(ys.len());
    if n < 2 { return None; }
    let mean_x = xs[..n].iter().sum::<f64>() / n as f64;
    let mean_y = ys[..n].iter().sum::<f64>() / n as f64;
    let mut ss_xy = 0.0;
    let mut ss_xx = 0.0;
    for i in 0..n {
        ss_xy += (xs[i] - mean_x) * (ys[i] - mean_y);
        ss_xx += (xs[i] - mean_x) * (xs[i] - mean_x);
    }
    if ss_xx == 0.0 { return None; }
    let slope = ss_xy / ss_xx;
    let intercept = mean_y - slope * mean_x;
    Some((slope, intercept, xs[..n].to_vec(), ys[..n].to_vec()))
}

fn fn_slope(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match linear_regression(args, ctx) {
        Some((slope, _, _, _)) => CellValue::Number(slope),
        None => CellValue::Error(CellError::Na),
    }
}

fn fn_intercept(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match linear_regression(args, ctx) {
        Some((_, intercept, _, _)) => CellValue::Number(intercept),
        None => CellValue::Error(CellError::Na),
    }
}

fn fn_rsq(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match linear_regression(args, ctx) {
        Some((slope, intercept, xs, ys)) => {
            let n = xs.len();
            let mean_y = ys.iter().sum::<f64>() / n as f64;
            let ss_tot: f64 = ys.iter().map(|y| (y - mean_y).powi(2)).sum();
            let ss_res: f64 = (0..n).map(|i| (ys[i] - (slope * xs[i] + intercept)).powi(2)).sum();
            if ss_tot == 0.0 { CellValue::Error(CellError::Div0) }
            else { CellValue::Number(1.0 - ss_res / ss_tot) }
        }
        None => CellValue::Error(CellError::Na),
    }
}

fn fn_forecast(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) => n,
        None => return CellValue::Error(CellError::Value),
    };
    let reg_args = [args[1].clone(), args[2].clone()];
    match linear_regression(&reg_args, ctx) {
        Some((slope, intercept, _, _)) => CellValue::Number(slope * x + intercept),
        None => CellValue::Error(CellError::Na),
    }
}

fn fn_steyx(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match linear_regression(args, ctx) {
        Some((slope, intercept, xs, ys)) => {
            let n = xs.len();
            if n < 3 { return CellValue::Error(CellError::Div0); }
            let ss_res: f64 = (0..n).map(|i| (ys[i] - (slope * xs[i] + intercept)).powi(2)).sum();
            CellValue::Number((ss_res / (n as f64 - 2.0)).sqrt())
        }
        None => CellValue::Error(CellError::Na),
    }
}

fn fn_skew(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    let n = nums.len();
    if n < 3 { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / n as f64;
    let s = (nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (n as f64 - 1.0)).sqrt();
    if s == 0.0 { return CellValue::Error(CellError::Div0); }
    let m3: f64 = nums.iter().map(|x| ((x - mean) / s).powi(3)).sum();
    CellValue::Number(n as f64 / ((n as f64 - 1.0) * (n as f64 - 2.0)) * m3)
}

fn fn_kurt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    let n = nums.len() as f64;
    if n < 4.0 { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / n;
    let s = (nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (n - 1.0)).sqrt();
    if s == 0.0 { return CellValue::Error(CellError::Div0); }
    let m4: f64 = nums.iter().map(|x| ((x - mean) / s).powi(4)).sum();
    let k = (n * (n + 1.0)) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) * m4
        - 3.0 * (n - 1.0).powi(2) / ((n - 2.0) * (n - 3.0));
    CellValue::Number(k)
}

fn collect_pair_arrays(args: &[Expr], ctx: &dyn EvalContext) -> Option<(Vec<f64>, Vec<f64>)> {
    let xs = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return None,
    };
    let ys = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return None,
    };
    let n = xs.len().min(ys.len());
    if n < 2 { return None; }
    Some((xs[..n].to_vec(), ys[..n].to_vec()))
}

fn fn_covariance_p(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match collect_pair_arrays(args, ctx) {
        Some((xs, ys)) => {
            let n = xs.len();
            let mean_x = xs.iter().sum::<f64>() / n as f64;
            let mean_y = ys.iter().sum::<f64>() / n as f64;
            let cov: f64 = (0..n).map(|i| (xs[i] - mean_x) * (ys[i] - mean_y)).sum::<f64>() / n as f64;
            CellValue::Number(cov)
        }
        None => CellValue::Error(CellError::Na),
    }
}

fn fn_covariance_s(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match collect_pair_arrays(args, ctx) {
        Some((xs, ys)) => {
            let n = xs.len();
            if n < 2 { return CellValue::Error(CellError::Div0); }
            let mean_x = xs.iter().sum::<f64>() / n as f64;
            let mean_y = ys.iter().sum::<f64>() / n as f64;
            let cov: f64 = (0..n).map(|i| (xs[i] - mean_x) * (ys[i] - mean_y)).sum::<f64>() / (n as f64 - 1.0);
            CellValue::Number(cov)
        }
        None => CellValue::Error(CellError::Na),
    }
}

fn fn_geomean(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() || nums.iter().any(|&n| n <= 0.0) {
        return CellValue::Error(CellError::Num);
    }
    let log_sum: f64 = nums.iter().map(|n| n.ln()).sum();
    CellValue::Number((log_sum / nums.len() as f64).exp())
}

fn fn_harmean(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() || nums.iter().any(|&n| n <= 0.0) {
        return CellValue::Error(CellError::Num);
    }
    let recip_sum: f64 = nums.iter().map(|n| 1.0 / n).sum();
    CellValue::Number(nums.len() as f64 / recip_sum)
}

fn fn_trimmean(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    let pct = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if (0.0..1.0).contains(&n) => n,
        _ => return CellValue::Error(CellError::Num),
    };
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let trim = ((nums.len() as f64 * pct) / 2.0).floor() as usize;
    let trimmed = &nums[trim..nums.len() - trim];
    if trimmed.is_empty() { return CellValue::Error(CellError::Num); }
    CellValue::Number(trimmed.iter().sum::<f64>() / trimmed.len() as f64)
}

fn fn_avedev(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let ad: f64 = nums.iter().map(|x| (x - mean).abs()).sum();
    CellValue::Number(ad / nums.len() as f64)
}

fn fn_devsq(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    CellValue::Number(nums.iter().map(|x| (x - mean).powi(2)).sum())
}

fn fn_standardize(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let mean = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sd = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if sd <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((x - mean) / sd)
}

fn fn_fisher(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match evaluate(&args[0], ctx, reg).as_f64() {
        Some(x) if (-1.0..1.0).contains(&x) => CellValue::Number(0.5 * ((1.0 + x) / (1.0 - x)).ln()),
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_fisherinv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match evaluate(&args[0], ctx, reg).as_f64() {
        Some(y) => {
            let e2y = (2.0 * y).exp();
            CellValue::Number((e2y - 1.0) / (e2y + 1.0))
        }
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_prob(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let xs = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    let probs = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    let lower = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let upper = if args.len() > 3 {
        match evaluate(&args[3], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) }
    } else { lower };
    let n = xs.len().min(probs.len());
    let mut total = 0.0;
    for i in 0..n {
        if xs[i] >= lower && xs[i] <= upper { total += probs[i]; }
    }
    CellValue::Number(total)
}

fn fn_permut(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let k = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    if k > n { return CellValue::Error(CellError::Num); }
    let mut result: f64 = 1.0;
    for i in 0..k { result *= (n - i) as f64; }
    CellValue::Number(result)
}

fn fn_permutationa(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let k = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    CellValue::Number(n.powf(k))
}

// ---------------------------------------------------------------------------
// Statistical distribution helper functions
// ---------------------------------------------------------------------------

/// Error function approximation (Abramowitz & Stegun, formula 7.1.26)
fn erf_approx(x: f64) -> f64 {
    let sign = if x >= 0.0 { 1.0 } else { -1.0 };
    let x = x.abs();
    let t = 1.0 / (1.0 + 0.3275911 * x);
    let poly = t * (0.254829592
        + t * (-0.284496736
        + t * (1.421413741
        + t * (-1.453152027
        + t * 1.061405429))));
    sign * (1.0 - poly * (-x * x).exp())
}

/// Standard normal PDF
fn std_normal_pdf(x: f64) -> f64 {
    (-0.5 * x * x).exp() / (2.0 * std::f64::consts::PI).sqrt()
}

/// Standard normal CDF
fn std_normal_cdf(x: f64) -> f64 {
    0.5 * (1.0 + erf_approx(x / std::f64::consts::SQRT_2))
}

/// Inverse standard normal CDF (rational approximation, Abramowitz & Stegun)
fn std_normal_inv(p: f64) -> f64 {
    if p <= 0.0 || p >= 1.0 {
        return f64::NAN;
    }
    if p < 0.5 {
        let t = (-2.0 * p.ln()).sqrt();
        let c0 = 2.515517;
        let c1 = 0.802853;
        let c2 = 0.010328;
        let d1 = 1.432788;
        let d2 = 0.189269;
        let d3 = 0.001308;
        -(t - (c0 + c1 * t + c2 * t * t) / (1.0 + d1 * t + d2 * t * t + d3 * t * t * t))
    } else {
        -std_normal_inv(1.0 - p)
    }
}

/// Lanczos approximation for the gamma function
fn gamma_fn(x: f64) -> f64 {
    if x <= 0.0 && x == x.floor() {
        return f64::INFINITY;
    }
    let g = 7.0;
    let c = [
        0.99999999999980993,
        676.5203681218851,
        -1259.1392167224028,
        771.32342877765313,
        -176.61502916214059,
        12.507343278686905,
        -0.13857109526572012,
        9.9843695780195716e-6,
        1.5056327351493116e-7,
    ];
    if x < 0.5 {
        std::f64::consts::PI / ((std::f64::consts::PI * x).sin() * gamma_fn(1.0 - x))
    } else {
        let x = x - 1.0;
        let mut sum = c[0];
        for (i, &ci) in c[1..].iter().enumerate() {
            sum += ci / (x + i as f64 + 1.0);
        }
        let t = x + g + 0.5;
        (2.0 * std::f64::consts::PI).sqrt() * t.powf(x + 0.5) * (-t).exp() * sum
    }
}

/// Natural log of gamma function
fn ln_gamma(x: f64) -> f64 {
    gamma_fn(x).abs().ln()
}

/// Regularized lower incomplete gamma function P(a, x) via series expansion
fn regularized_gamma_p(a: f64, x: f64) -> f64 {
    if x < 0.0 {
        return 0.0;
    }
    if x == 0.0 {
        return 0.0;
    }
    // Use series expansion for x < a + 1, continued fraction otherwise
    if x < a + 1.0 {
        // Series: P(a,x) = e^(-x) * x^a * sum(x^n / gamma(a+n+1))
        let mut sum = 1.0 / a;
        let mut term = 1.0 / a;
        for n in 1..200 {
            term *= x / (a + n as f64);
            sum += term;
            if term.abs() < 1e-14 * sum.abs() {
                break;
            }
        }
        sum * (-x + a * x.ln() - ln_gamma(a)).exp()
    } else {
        // Continued fraction (Lentz's method) for upper incomplete gamma
        1.0 - regularized_gamma_q_cf(a, x)
    }
}

/// Regularized upper incomplete gamma Q(a,x) via continued fraction
fn regularized_gamma_q_cf(a: f64, x: f64) -> f64 {
    let mut c = 1e-30_f64;
    let mut d = 1.0 / (x + 1.0 - a);
    let mut f = d;
    for n in 1..200 {
        let an = n as f64 * (a - n as f64);
        let bn = x + 2.0 * n as f64 + 1.0 - a;
        d = bn + an * d;
        if d.abs() < 1e-30 { d = 1e-30; }
        d = 1.0 / d;
        c = bn + an / c;
        if c.abs() < 1e-30 { c = 1e-30; }
        let delta = c * d;
        f *= delta;
        if (delta - 1.0).abs() < 1e-14 {
            break;
        }
    }
    f * (-x + a * x.ln() - ln_gamma(a)).exp()
}

/// Regularized incomplete beta function I_x(a, b) via continued fraction
fn regularized_beta(x: f64, a: f64, b: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    if x >= 1.0 {
        return 1.0;
    }
    // Use symmetry relation when x > (a+1)/(a+b+2)
    if x > (a + 1.0) / (a + b + 2.0) {
        return 1.0 - regularized_beta(1.0 - x, b, a);
    }
    let lbeta = ln_gamma(a) + ln_gamma(b) - ln_gamma(a + b);
    let front = (a * x.ln() + b * (1.0 - x).ln() - lbeta).exp() / a;
    // Continued fraction (Lentz's method)
    let mut c = 1e-30_f64;
    let mut d = 1.0 / (1.0 - (a + b) * x / (a + 1.0));
    if d.abs() < 1e-30 { d = 1e-30; }
    let mut f = d;
    for m in 1..200 {
        // Even step
        let m_f = m as f64;
        let num_even = m_f * (b - m_f) * x / ((a + 2.0 * m_f - 1.0) * (a + 2.0 * m_f));
        d = 1.0 + num_even * d;
        if d.abs() < 1e-30 { d = 1e-30; }
        d = 1.0 / d;
        c = 1.0 + num_even / c;
        if c.abs() < 1e-30 { c = 1e-30; }
        f *= c * d;

        // Odd step
        let num_odd = -((a + m_f) * (a + b + m_f) * x) / ((a + 2.0 * m_f) * (a + 2.0 * m_f + 1.0));
        d = 1.0 + num_odd * d;
        if d.abs() < 1e-30 { d = 1e-30; }
        d = 1.0 / d;
        c = 1.0 + num_odd / c;
        if c.abs() < 1e-30 { c = 1e-30; }
        let delta = c * d;
        f *= delta;
        if (delta - 1.0).abs() < 1e-14 {
            break;
        }
    }
    front * f * a
}

/// T-distribution CDF using incomplete beta
fn t_dist_cdf(x: f64, df: f64) -> f64 {
    let xt = df / (df + x * x);
    0.5 * (1.0 + (1.0 - regularized_beta(xt, df / 2.0, 0.5)).copysign(x))
}

/// T-distribution PDF
fn t_dist_pdf(x: f64, df: f64) -> f64 {
    let coeff = gamma_fn((df + 1.0) / 2.0) / (gamma_fn(df / 2.0) * (df * std::f64::consts::PI).sqrt());
    coeff * (1.0 + x * x / df).powf(-(df + 1.0) / 2.0)
}

/// Chi-squared CDF: P(k/2, x/2)
fn chisq_cdf(x: f64, df: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    regularized_gamma_p(df / 2.0, x / 2.0)
}

/// Chi-squared PDF
fn chisq_pdf(x: f64, df: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    let k2 = df / 2.0;
    x.powf(k2 - 1.0) * (-x / 2.0).exp() / (2.0_f64.powf(k2) * gamma_fn(k2))
}

/// F-distribution CDF using incomplete beta
fn f_dist_cdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    regularized_beta(d1 * x / (d1 * x + d2), d1 / 2.0, d2 / 2.0)
}

/// F-distribution PDF
fn f_dist_pdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    let num = ((d1 * x).powf(d1) * d2.powf(d2) / (d1 * x + d2).powf(d1 + d2)).sqrt();
    num / (x * beta_fn(d1 / 2.0, d2 / 2.0))
}

/// Beta function B(a,b) = Gamma(a)*Gamma(b)/Gamma(a+b)
fn beta_fn(a: f64, b: f64) -> f64 {
    (ln_gamma(a) + ln_gamma(b) - ln_gamma(a + b)).exp()
}

/// Newton-Raphson inversion: find x such that cdf(x) = p
fn newton_raphson_inv(
    p: f64,
    initial: f64,
    cdf: impl Fn(f64) -> f64,
    pdf: impl Fn(f64) -> f64,
    iterations: usize,
) -> f64 {
    let mut x = initial;
    for _ in 0..iterations {
        let fx = cdf(x) - p;
        let dfx = pdf(x);
        if dfx.abs() < 1e-30 {
            break;
        }
        x -= fx / dfx;
    }
    x
}

/// Binomial coefficient C(n, k)
fn binom_coeff(n: u64, k: u64) -> f64 {
    if k > n {
        return 0.0;
    }
    let k = k.min(n - k);
    let mut result = 1.0;
    for i in 0..k {
        result *= (n - i) as f64 / (i + 1) as f64;
    }
    result
}

// ---------------------------------------------------------------------------
// Distribution function implementations
// ---------------------------------------------------------------------------

fn fn_norm_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let mean = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sd = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[3], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if sd <= 0.0 { return CellValue::Error(CellError::Num); }
    let z = (x - mean) / sd;
    if cumulative {
        CellValue::Number(std_normal_cdf(z))
    } else {
        CellValue::Number(std_normal_pdf(z) / sd)
    }
}

fn fn_norm_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let mean = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sd = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p <= 0.0 || p >= 1.0 || sd <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(mean + sd * std_normal_inv(p))
}

fn fn_norm_s_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let z = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if cumulative {
        CellValue::Number(std_normal_cdf(z))
    } else {
        CellValue::Number(std_normal_pdf(z))
    }
}

fn fn_norm_s_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p <= 0.0 || p >= 1.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(std_normal_inv(p))
}

fn fn_t_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if df < 1.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        CellValue::Number(t_dist_cdf(x, df))
    } else {
        CellValue::Number(t_dist_pdf(x, df))
    }
}

fn fn_t_dist_2t(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if x < 0.0 || df < 1.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(2.0 * (1.0 - t_dist_cdf(x.abs(), df)))
}

fn fn_t_dist_rt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if df < 1.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(1.0 - t_dist_cdf(x, df))
}

fn fn_t_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p <= 0.0 || p >= 1.0 || df < 1.0 { return CellValue::Error(CellError::Num); }
    let initial = std_normal_inv(p);
    let result = newton_raphson_inv(
        p, initial,
        |x| t_dist_cdf(x, df),
        |x| t_dist_pdf(x, df),
        50,
    );
    CellValue::Number(result)
}

fn fn_t_inv_2t(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p <= 0.0 || p >= 1.0 || df < 1.0 { return CellValue::Error(CellError::Num); }
    // T.INV.2T(p, df) = T.INV(1 - p/2, df) but positive
    let target = 1.0 - p / 2.0;
    let initial = std_normal_inv(target).abs();
    let result = newton_raphson_inv(
        target, initial,
        |x| t_dist_cdf(x, df),
        |x| t_dist_pdf(x, df),
        50,
    );
    CellValue::Number(result.abs())
}

fn fn_binom_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let k = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let n = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let p = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[3], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if p < 0.0 || p > 1.0 || k > n { return CellValue::Error(CellError::Num); }
    if cumulative {
        let mut sum = 0.0;
        for i in 0..=k {
            sum += binom_coeff(n, i) * p.powi(i as i32) * (1.0 - p).powi((n - i) as i32);
        }
        CellValue::Number(sum)
    } else {
        CellValue::Number(binom_coeff(n, k) * p.powi(k as i32) * (1.0 - p).powi((n - k) as i32))
    }
}

fn fn_binom_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let p = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let alpha = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if p < 0.0 || p > 1.0 || alpha < 0.0 || alpha > 1.0 { return CellValue::Error(CellError::Num); }
    let mut cum = 0.0;
    for k in 0..=n {
        cum += binom_coeff(n, k) * p.powi(k as i32) * (1.0 - p).powi((n - k) as i32);
        if cum >= alpha {
            return CellValue::Number(k as f64);
        }
    }
    CellValue::Number(n as f64)
}

fn fn_poisson_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let mean = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if mean < 0.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        let mut sum = 0.0;
        let mut term = (-mean).exp(); // k=0 term
        sum += term;
        for k in 1..=x {
            term *= mean / k as f64;
            sum += term;
        }
        CellValue::Number(sum)
    } else {
        // e^(-mean) * mean^x / x!
        let log_pmf = -mean + x as f64 * mean.ln() - ln_gamma(x as f64 + 1.0);
        CellValue::Number(log_pmf.exp())
    }
}

fn fn_expon_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let lambda = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if lambda <= 0.0 || x < 0.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        CellValue::Number(1.0 - (-lambda * x).exp())
    } else {
        CellValue::Number(lambda * (-lambda * x).exp())
    }
}

fn fn_gamma(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if x <= 0.0 && x == x.floor() { return CellValue::Error(CellError::Num); }
    CellValue::Number(gamma_fn(x))
}

fn fn_gammaln(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if x <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(ln_gamma(x))
}

fn fn_gamma_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let alpha = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let beta = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[3], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if alpha <= 0.0 || beta <= 0.0 || x < 0.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        CellValue::Number(regularized_gamma_p(alpha, x / beta))
    } else {
        // PDF: x^(alpha-1) * exp(-x/beta) / (beta^alpha * Gamma(alpha))
        let log_pdf = (alpha - 1.0) * x.ln() - x / beta - alpha * beta.ln() - ln_gamma(alpha);
        CellValue::Number(log_pdf.exp())
    }
}

fn fn_gamma_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let alpha = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let beta = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p < 0.0 || p > 1.0 || alpha <= 0.0 || beta <= 0.0 { return CellValue::Error(CellError::Num); }
    if p == 0.0 { return CellValue::Number(0.0); }
    if p == 1.0 { return CellValue::Error(CellError::Num); }
    // Newton-Raphson on the gamma CDF
    let mut x = alpha * beta; // initial guess at the mean
    for _ in 0..50 {
        let cdf_val = regularized_gamma_p(alpha, x / beta);
        let log_pdf = (alpha - 1.0) * x.ln() - x / beta - alpha * beta.ln() - ln_gamma(alpha);
        let pdf_val = log_pdf.exp();
        if pdf_val.abs() < 1e-30 { break; }
        x -= (cdf_val - p) / pdf_val;
        if x <= 0.0 { x = 1e-10; }
    }
    CellValue::Number(x)
}

fn fn_beta_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let alpha = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let beta = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[3], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    let a = if args.len() > 4 { evaluate(&args[4], ctx, reg).as_f64().unwrap_or(0.0) } else { 0.0 };
    let b = if args.len() > 5 { evaluate(&args[5], ctx, reg).as_f64().unwrap_or(1.0) } else { 1.0 };
    if alpha <= 0.0 || beta <= 0.0 || a >= b { return CellValue::Error(CellError::Num); }
    let z = (x - a) / (b - a);
    if z < 0.0 || z > 1.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        CellValue::Number(regularized_beta(z, alpha, beta))
    } else {
        // PDF: (x-a)^(alpha-1) * (b-x)^(beta-1) / ((b-a)^(alpha+beta-1) * B(alpha,beta))
        let log_pdf = (alpha - 1.0) * z.ln() + (beta - 1.0) * (1.0 - z).ln()
            - (ln_gamma(alpha) + ln_gamma(beta) - ln_gamma(alpha + beta))
            - (b - a).ln();
        CellValue::Number(log_pdf.exp())
    }
}

fn fn_beta_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let alpha = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let beta = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let a = if args.len() > 3 { evaluate(&args[3], ctx, reg).as_f64().unwrap_or(0.0) } else { 0.0 };
    let b = if args.len() > 4 { evaluate(&args[4], ctx, reg).as_f64().unwrap_or(1.0) } else { 1.0 };
    if p < 0.0 || p > 1.0 || alpha <= 0.0 || beta <= 0.0 || a >= b { return CellValue::Error(CellError::Num); }
    if p == 0.0 { return CellValue::Number(a); }
    if p == 1.0 { return CellValue::Number(b); }
    // Newton-Raphson on regularized_beta
    let mut x = alpha / (alpha + beta); // initial guess
    for _ in 0..50 {
        let cdf_val = regularized_beta(x, alpha, beta);
        let log_pdf = (alpha - 1.0) * x.ln() + (beta - 1.0) * (1.0 - x).ln()
            - (ln_gamma(alpha) + ln_gamma(beta) - ln_gamma(alpha + beta));
        let pdf_val = log_pdf.exp();
        if pdf_val.abs() < 1e-30 { break; }
        x -= (cdf_val - p) / pdf_val;
        x = x.clamp(1e-10, 1.0 - 1e-10);
    }
    CellValue::Number(a + x * (b - a))
}

fn fn_weibull_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let alpha = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let beta = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[3], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if alpha <= 0.0 || beta <= 0.0 || x < 0.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        CellValue::Number(1.0 - (-(x / beta).powf(alpha)).exp())
    } else {
        CellValue::Number((alpha / beta) * (x / beta).powf(alpha - 1.0) * (-(x / beta).powf(alpha)).exp())
    }
}

fn fn_lognorm_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let mean = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sd = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[3], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if sd <= 0.0 || x <= 0.0 { return CellValue::Error(CellError::Num); }
    let z = (x.ln() - mean) / sd;
    if cumulative {
        CellValue::Number(std_normal_cdf(z))
    } else {
        CellValue::Number(std_normal_pdf(z) / (x * sd))
    }
}

fn fn_lognorm_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let mean = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sd = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p <= 0.0 || p >= 1.0 || sd <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((mean + sd * std_normal_inv(p)).exp())
}

fn fn_chisq_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if df < 1.0 || x < 0.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        CellValue::Number(chisq_cdf(x, df))
    } else {
        CellValue::Number(chisq_pdf(x, df))
    }
}

fn fn_chisq_dist_rt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if df < 1.0 || x < 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(1.0 - chisq_cdf(x, df))
}

fn fn_chisq_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p < 0.0 || p > 1.0 || df < 1.0 { return CellValue::Error(CellError::Num); }
    if p == 0.0 { return CellValue::Number(0.0); }
    if p == 1.0 { return CellValue::Error(CellError::Num); }
    let mut x = df; // initial guess at mean
    for _ in 0..50 {
        let cdf_val = chisq_cdf(x, df);
        let pdf_val = chisq_pdf(x, df);
        if pdf_val.abs() < 1e-30 { break; }
        x -= (cdf_val - p) / pdf_val;
        if x <= 0.0 { x = 1e-10; }
    }
    CellValue::Number(x)
}

fn fn_chisq_inv_rt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p < 0.0 || p > 1.0 || df < 1.0 { return CellValue::Error(CellError::Num); }
    if p == 1.0 { return CellValue::Number(0.0); }
    if p == 0.0 { return CellValue::Error(CellError::Num); }
    // Invert 1 - chisq_cdf(x, df) = p => chisq_cdf(x, df) = 1 - p
    let target = 1.0 - p;
    let mut x = df;
    for _ in 0..50 {
        let cdf_val = chisq_cdf(x, df);
        let pdf_val = chisq_pdf(x, df);
        if pdf_val.abs() < 1e-30 { break; }
        x -= (cdf_val - target) / pdf_val;
        if x <= 0.0 { x = 1e-10; }
    }
    CellValue::Number(x)
}

fn fn_f_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d1 = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d2 = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[3], ctx, reg).as_f64() { Some(n) => n != 0.0, None => return CellValue::Error(CellError::Value) };
    if d1 < 1.0 || d2 < 1.0 || x < 0.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        CellValue::Number(f_dist_cdf(x, d1, d2))
    } else {
        CellValue::Number(f_dist_pdf(x, d1, d2))
    }
}

fn fn_f_dist_rt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d1 = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d2 = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if d1 < 1.0 || d2 < 1.0 || x < 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(1.0 - f_dist_cdf(x, d1, d2))
}

fn fn_f_inv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d1 = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d2 = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p < 0.0 || p > 1.0 || d1 < 1.0 || d2 < 1.0 { return CellValue::Error(CellError::Num); }
    if p == 0.0 { return CellValue::Number(0.0); }
    if p == 1.0 { return CellValue::Error(CellError::Num); }
    let mut x = d2 / (d2 - 2.0).max(0.1); // initial guess near mean
    for _ in 0..50 {
        let cdf_val = f_dist_cdf(x, d1, d2);
        let pdf_val = f_dist_pdf(x, d1, d2);
        if pdf_val.abs() < 1e-30 { break; }
        x -= (cdf_val - p) / pdf_val;
        if x <= 0.0 { x = 1e-10; }
    }
    CellValue::Number(x)
}

fn fn_f_inv_rt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let p = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d1 = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d2 = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if p < 0.0 || p > 1.0 || d1 < 1.0 || d2 < 1.0 { return CellValue::Error(CellError::Num); }
    if p == 1.0 { return CellValue::Number(0.0); }
    if p == 0.0 { return CellValue::Error(CellError::Num); }
    let target = 1.0 - p;
    let mut x = d2 / (d2 - 2.0).max(0.1);
    for _ in 0..50 {
        let cdf_val = f_dist_cdf(x, d1, d2);
        let pdf_val = f_dist_pdf(x, d1, d2);
        if pdf_val.abs() < 1e-30 { break; }
        x -= (cdf_val - target) / pdf_val;
        if x <= 0.0 { x = 1e-10; }
    }
    CellValue::Number(x)
}

fn fn_confidence_norm(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let alpha = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sd = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let size = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || size < 1.0 { return CellValue::Error(CellError::Num); }
    let z = std_normal_inv(1.0 - alpha / 2.0);
    CellValue::Number(z * sd / size.sqrt())
}

fn fn_confidence_t(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let alpha = match evaluate(&args[0], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sd = match evaluate(&args[1], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let size = match evaluate(&args[2], ctx, reg).as_f64() { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || size < 2.0 { return CellValue::Error(CellError::Num); }
    let df = size - 1.0;
    // Find t critical value via Newton-Raphson on t CDF
    let target = 1.0 - alpha / 2.0;
    let initial = std_normal_inv(target);
    let t_crit = newton_raphson_inv(
        target, initial,
        |x| t_dist_cdf(x, df),
        |x| t_dist_pdf(x, df),
        50,
    );
    CellValue::Number(t_crit * sd / size.sqrt())
}

fn fn_rank_avg(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let number = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) => n,
        None => return CellValue::Error(CellError::Value),
    };
    let nums = match &args[1] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let descending = if args.len() > 2 {
        evaluate(&args[2], ctx, reg).as_f64().unwrap_or(0.0) == 0.0
    } else {
        true
    };

    let (better, tied) = if descending {
        (
            nums.iter().filter(|&&n| n > number).count(),
            nums.iter().filter(|&&n| (n - number).abs() < f64::EPSILON).count(),
        )
    } else {
        (
            nums.iter().filter(|&&n| n < number).count(),
            nums.iter().filter(|&&n| (n - number).abs() < f64::EPSILON).count(),
        )
    };
    if tied == 0 {
        return CellValue::Error(CellError::Na);
    }
    // Average rank of tied positions: (better+1 + better+tied) / 2
    let avg_rank = (better as f64 + 1.0 + better as f64 + tied as f64) / 2.0;
    CellValue::Number(avg_rank)
}

fn fn_quartile_exc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    let q = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if (1.0..=3.0).contains(&n) && n == n.floor() => n as usize,
        _ => return CellValue::Error(CellError::Num),
    };
    if nums.is_empty() { return CellValue::Error(CellError::Num); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    // Exclusive percentile at q/4
    let k = q as f64 / 4.0;
    let n = nums.len();
    if k <= 1.0 / (n as f64 + 1.0) || k >= n as f64 / (n as f64 + 1.0) {
        return CellValue::Error(CellError::Num);
    }
    let rank = k * (n as f64 + 1.0) - 1.0;
    let lower = rank.floor() as usize;
    let upper = (lower + 1).min(n - 1);
    let frac = rank - lower as f64;
    CellValue::Number(nums[lower] * (1.0 - frac) + nums[upper] * frac)
}

fn fn_skew_p(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    let n = nums.len();
    if n < 3 { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / n as f64;
    let s = (nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / n as f64).sqrt();
    if s == 0.0 { return CellValue::Error(CellError::Div0); }
    let m3: f64 = nums.iter().map(|x| ((x - mean) / s).powi(3)).sum();
    CellValue::Number(m3 / n as f64)
}

fn fn_growth(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_trend(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_combin_stat(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let k = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    if k > n { return CellValue::Error(CellError::Num); }
    CellValue::Number(binom_coeff(n, k).round())
}

fn fn_combina_stat(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let k = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let total = n + k - 1;
    if k > total { return CellValue::Error(CellError::Num); }
    CellValue::Number(binom_coeff(total, k).round())
}

fn fn_hypgeom_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let n = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let big_k = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let big_n = match evaluate(&args[3], ctx, reg).as_f64() { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[4], ctx, reg).as_f64() { Some(v) => v != 0.0, None => return CellValue::Error(CellError::Value) };
    if n > big_n || big_k > big_n || s > n || s > big_k { return CellValue::Error(CellError::Num); }
    let pmf = |x: u64| -> f64 {
        binom_coeff(big_k, x) * binom_coeff(big_n - big_k, n - x) / binom_coeff(big_n, n)
    };
    if cumulative {
        let mut sum = 0.0;
        let lo = if n > big_n - big_k { n - (big_n - big_k) } else { 0 };
        for i in lo..=s {
            sum += pmf(i);
        }
        CellValue::Number(sum)
    } else {
        CellValue::Number(pmf(s))
    }
}

fn fn_z_test(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return CellValue::Error(CellError::Value),
    };
    let x0 = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let n = nums.len();
    if n == 0 { return CellValue::Error(CellError::Na); }
    let mean = nums.iter().sum::<f64>() / n as f64;
    let sigma = if args.len() > 2 {
        match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) }
    } else {
        // sample standard deviation
        if n < 2 { return CellValue::Error(CellError::Na); }
        (nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (n as f64 - 1.0)).sqrt()
    };
    if sigma <= 0.0 { return CellValue::Error(CellError::Div0); }
    let z = (mean - x0) / (sigma / (n as f64).sqrt());
    CellValue::Number(1.0 - std_normal_cdf(z))
}

fn fn_betadist_compat(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let alpha = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let beta_val = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let a = if args.len() > 3 { evaluate(&args[3], ctx, reg).as_f64().unwrap_or(0.0) } else { 0.0 };
    let b = if args.len() > 4 { evaluate(&args[4], ctx, reg).as_f64().unwrap_or(1.0) } else { 1.0 };
    if b <= a || alpha <= 0.0 || beta_val <= 0.0 { return CellValue::Error(CellError::Num); }
    let normalized = (x - a) / (b - a);
    CellValue::Number(regularized_beta(normalized, alpha, beta_val))
}

fn fn_binom_dist_range(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let trials = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let prob = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let s1 = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let s2 = if args.len() > 3 { evaluate(&args[3], ctx, reg).as_f64().unwrap_or(s1 as f64) as u64 } else { s1 };
    let mut total = 0.0;
    for s in s1..=s2 {
        let c = binom_coeff(trials, s);
        total += c * prob.powi(s as i32) * (1.0 - prob).powi((trials - s) as i32);
    }
    CellValue::Number(total)
}

fn fn_chisq_test(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let actual = match &args[0] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) };
    let expected = match &args[1] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) };
    if actual.len() != expected.len() || actual.is_empty() { return CellValue::Error(CellError::Na); }
    let mut chi2 = 0.0;
    let mut df = 0usize;
    for (a, e) in actual.iter().zip(expected.iter()) {
        if let (Some(av), Some(ev)) = (a.as_f64(), e.as_f64()) {
            if ev == 0.0 { return CellValue::Error(CellError::Div0); }
            chi2 += (av - ev).powi(2) / ev;
            df += 1;
        }
    }
    if df <= 1 { return CellValue::Error(CellError::Na); }
    let k = (df - 1) as f64;
    CellValue::Number(1.0 - chisq_cdf(chi2, k))
}

fn fn_f_test(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let arr1 = match &args[0] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) };
    let arr2 = match &args[1] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) };
    let nums1: Vec<f64> = arr1.iter().filter_map(|v| v.as_f64()).collect();
    let nums2: Vec<f64> = arr2.iter().filter_map(|v| v.as_f64()).collect();
    if nums1.len() < 2 || nums2.len() < 2 { return CellValue::Error(CellError::Na); }
    let var1 = match variance(&nums1, false) { Some(v) => v, None => return CellValue::Error(CellError::Na) };
    let var2 = match variance(&nums2, false) { Some(v) => v, None => return CellValue::Error(CellError::Na) };
    if var2 == 0.0 { return CellValue::Error(CellError::Div0); }
    let f = var1 / var2;
    let d1 = (nums1.len() - 1) as f64;
    let d2 = (nums2.len() - 1) as f64;
    let p = 1.0 - f_dist_cdf(f, d1, d2);
    CellValue::Number(2.0 * p.min(1.0 - p))
}

fn fn_gauss(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let z = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    CellValue::Number(std_normal_cdf(z) - 0.5)
}

fn fn_hypgeomdist_compat(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let n = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let kk = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let nn = match evaluate(&args[3], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let p = binom_coeff(kk, s) * binom_coeff(nn - kk, n - s) / binom_coeff(nn, n);
    CellValue::Number(p)
}

fn fn_linest(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let ys = match &args[0] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) };
    let xs = if args.len() > 1 {
        match &args[1] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) }
    } else {
        (1..=ys.len()).map(|i| CellValue::Number(i as f64)).collect()
    };
    let y_nums: Vec<f64> = ys.iter().filter_map(|v| v.as_f64()).collect();
    let x_nums: Vec<f64> = xs.iter().filter_map(|v| v.as_f64()).collect();
    let n = y_nums.len().min(x_nums.len());
    if n < 2 { return CellValue::Error(CellError::Na); }
    let x_mean = x_nums[..n].iter().sum::<f64>() / n as f64;
    let y_mean = y_nums[..n].iter().sum::<f64>() / n as f64;
    let mut num = 0.0;
    let mut den = 0.0;
    for i in 0..n {
        num += (x_nums[i] - x_mean) * (y_nums[i] - y_mean);
        den += (x_nums[i] - x_mean).powi(2);
    }
    if den == 0.0 { return CellValue::Error(CellError::Na); }
    let slope = num / den;
    let intercept = y_mean - slope * x_mean;
    CellValue::Array(Box::new(vec![vec![CellValue::Number(slope), CellValue::Number(intercept)]]))
}

fn fn_logest(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let ys = match &args[0] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) };
    let xs = if args.len() > 1 {
        match &args[1] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) }
    } else {
        (1..=ys.len()).map(|i| CellValue::Number(i as f64)).collect()
    };
    let y_nums: Vec<f64> = ys.iter().filter_map(|v| v.as_f64()).collect();
    let x_nums: Vec<f64> = xs.iter().filter_map(|v| v.as_f64()).collect();
    let n = y_nums.len().min(x_nums.len());
    if n < 2 { return CellValue::Error(CellError::Na); }
    let log_ys: Vec<f64> = y_nums[..n].iter().map(|&y| if y > 0.0 { y.ln() } else { f64::NAN }).collect();
    if log_ys.iter().any(|v| v.is_nan()) { return CellValue::Error(CellError::Num); }
    let x_mean = x_nums[..n].iter().sum::<f64>() / n as f64;
    let y_mean = log_ys.iter().sum::<f64>() / n as f64;
    let mut num = 0.0;
    let mut den = 0.0;
    for i in 0..n {
        num += (x_nums[i] - x_mean) * (log_ys[i] - y_mean);
        den += (x_nums[i] - x_mean).powi(2);
    }
    if den == 0.0 { return CellValue::Error(CellError::Na); }
    let slope = num / den;
    let intercept = y_mean - slope * x_mean;
    CellValue::Array(Box::new(vec![vec![CellValue::Number(slope.exp()), CellValue::Number(intercept.exp())]]))
}

fn fn_lognormdist_compat(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let mean = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let std = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if x <= 0.0 || std <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(std_normal_cdf((x.ln() - mean) / std))
}

fn collect_all_numeric(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Vec<f64> {
    let mut nums = Vec::new();
    for arg in args {
        match arg {
            Expr::Range { start, end } => {
                for val in collect_range_values(start, end, ctx) {
                    match &val {
                        CellValue::Number(n) => nums.push(*n),
                        CellValue::Boolean(b) => nums.push(if *b { 1.0 } else { 0.0 }),
                        CellValue::String(s) => { if let Ok(n) = s.parse::<f64>() { nums.push(n); } else { nums.push(0.0); } }
                        CellValue::Empty => {}
                        _ => {}
                    }
                }
            }
            _ => {
                let val = evaluate(arg, ctx, reg);
                match &val {
                    CellValue::Number(n) => nums.push(*n),
                    CellValue::Boolean(b) => nums.push(if *b { 1.0 } else { 0.0 }),
                    CellValue::String(s) => { if let Ok(n) = s.parse::<f64>() { nums.push(n); } else { nums.push(0.0); } }
                    _ => {}
                }
            }
        }
    }
    nums
}

fn fn_averagea(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_all_numeric(args, ctx, reg);
    if nums.is_empty() { return CellValue::Error(CellError::Div0); }
    CellValue::Number(nums.iter().sum::<f64>() / nums.len() as f64)
}

fn fn_maxa(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_all_numeric(args, ctx, reg);
    if nums.is_empty() { return CellValue::Number(0.0); }
    CellValue::Number(nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max))
}

fn fn_mina(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_all_numeric(args, ctx, reg);
    if nums.is_empty() { return CellValue::Number(0.0); }
    CellValue::Number(nums.iter().cloned().fold(f64::INFINITY, f64::min))
}

fn fn_mode_mult(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() { return CellValue::Error(CellError::Na); }
    let mut counts = std::collections::HashMap::<u64, usize>::new();
    for &n in &nums {
        *counts.entry(n.to_bits()).or_insert(0) += 1;
    }
    let max_count = counts.values().cloned().max().unwrap_or(0);
    if max_count < 2 { return CellValue::Error(CellError::Na); }
    let modes: Vec<CellValue> = nums.iter().filter(|&&n| counts[&n.to_bits()] == max_count)
        .map(|&n| CellValue::Number(n))
        .collect::<Vec<_>>();
    let mut seen = std::collections::HashSet::new();
    let unique_modes: Vec<Vec<CellValue>> = modes.into_iter()
        .filter(|v| { if let CellValue::Number(n) = v { seen.insert(n.to_bits()) } else { true } })
        .map(|v| vec![v])
        .collect();
    CellValue::Array(Box::new(unique_modes))
}

fn fn_negbinom_dist(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let f = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let s = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let p = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let cumulative = match evaluate(&args[3], ctx, reg) { CellValue::Boolean(b) => b, CellValue::Number(n) => n != 0.0, _ => false };
    if p <= 0.0 || p >= 1.0 { return CellValue::Error(CellError::Num); }
    if cumulative {
        let mut total = 0.0;
        for i in 0..=f {
            total += binom_coeff(i + s - 1, i) * p.powi(s as i32) * (1.0 - p).powi(i as i32);
        }
        CellValue::Number(total)
    } else {
        let prob = binom_coeff(f + s - 1, f) * p.powi(s as i32) * (1.0 - p).powi(f as i32);
        CellValue::Number(prob)
    }
}

fn fn_negbinomdist_compat(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let f = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let s = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let p = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if p <= 0.0 || p >= 1.0 { return CellValue::Error(CellError::Num); }
    let prob = binom_coeff(f + s - 1, f) * p.powi(s as i32) * (1.0 - p).powi(f as i32);
    CellValue::Number(prob)
}

fn fn_normsdist_compat(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let z = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    CellValue::Number(std_normal_cdf(z))
}

fn fn_percentrank(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = match &args[0] { Expr::Range { start, end } => { let vals = collect_range_values(start, end, ctx); vals.iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>() }, _ => return CellValue::Error(CellError::Value) };
    let x = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let sig = if args.len() > 2 { evaluate(&args[2], ctx, reg).as_f64().unwrap_or(3.0) as usize } else { 3 };
    if nums.is_empty() { return CellValue::Error(CellError::Na); }
    let mut sorted = nums.clone();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let smaller = sorted.iter().filter(|&&v| v < x).count() as f64;
    let equal = sorted.iter().filter(|&&v| v == x).count() as f64;
    let n = sorted.len() as f64;
    let rank = (smaller + equal * 0.5 - 0.5) / (n - 1.0);
    let factor = 10f64.powi(sig as i32);
    CellValue::Number((rank * factor).floor() / factor)
}

fn fn_percentrank_exc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = match &args[0] { Expr::Range { start, end } => { let vals = collect_range_values(start, end, ctx); vals.iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>() }, _ => return CellValue::Error(CellError::Value) };
    let x = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let sig = if args.len() > 2 { evaluate(&args[2], ctx, reg).as_f64().unwrap_or(3.0) as usize } else { 3 };
    if nums.is_empty() { return CellValue::Error(CellError::Na); }
    let mut sorted = nums.clone();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let smaller = sorted.iter().filter(|&&v| v < x).count() as f64;
    let n = sorted.len() as f64;
    let rank = (smaller + 1.0) / (n + 1.0);
    let factor = 10f64.powi(sig as i32);
    CellValue::Number((rank * factor).floor() / factor)
}

fn fn_phi(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    CellValue::Number(std_normal_pdf(x))
}

fn fn_stdeva(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_all_numeric(args, ctx, reg);
    if nums.len() < 2 { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64;
    CellValue::Number(var.sqrt())
}

fn fn_stdevpa(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_all_numeric(args, ctx, reg);
    if nums.is_empty() { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / nums.len() as f64;
    CellValue::Number(var.sqrt())
}

fn fn_tdist_compat(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let df = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let tails = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let p = 1.0 - t_dist_cdf(x.abs(), df);
    match tails {
        1 => CellValue::Number(p),
        2 => CellValue::Number(2.0 * p),
        _ => CellValue::Error(CellError::Num),
    }
}

fn fn_t_test(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let arr1 = match &args[0] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) };
    let arr2 = match &args[1] { Expr::Range { start, end } => collect_range_values(start, end, ctx), _ => return CellValue::Error(CellError::Value) };
    let tails = match evaluate(&args[2], ctx, reg).as_f64() { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let _test_type = match evaluate(&args[3], ctx, reg).as_f64() { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let nums1: Vec<f64> = arr1.iter().filter_map(|v| v.as_f64()).collect();
    let nums2: Vec<f64> = arr2.iter().filter_map(|v| v.as_f64()).collect();
    let n1 = nums1.len() as f64;
    let n2 = nums2.len() as f64;
    if n1 < 2.0 || n2 < 2.0 { return CellValue::Error(CellError::Na); }
    let mean1 = nums1.iter().sum::<f64>() / n1;
    let mean2 = nums2.iter().sum::<f64>() / n2;
    let var1 = nums1.iter().map(|x| (x - mean1).powi(2)).sum::<f64>() / (n1 - 1.0);
    let var2 = nums2.iter().map(|x| (x - mean2).powi(2)).sum::<f64>() / (n2 - 1.0);
    let se = (var1 / n1 + var2 / n2).sqrt();
    if se == 0.0 { return CellValue::Error(CellError::Div0); }
    let t = (mean1 - mean2).abs() / se;
    let df = n1 + n2 - 2.0;
    let p = 1.0 - t_dist_cdf(t, df);
    match tails {
        1 => CellValue::Number(p),
        2 => CellValue::Number(2.0 * p),
        _ => CellValue::Error(CellError::Num),
    }
}

fn fn_vara(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_all_numeric(args, ctx, reg);
    if nums.len() < 2 { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    CellValue::Number(nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64)
}

fn fn_varpa(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_all_numeric(args, ctx, reg);
    if nums.is_empty() { return CellValue::Error(CellError::Div0); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    CellValue::Number(nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / nums.len() as f64)
}

fn fn_forecast_ets_stub(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}
