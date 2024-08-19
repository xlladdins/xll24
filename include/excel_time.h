// excel_time.h - Excel Julian date to time_t conversion
// Excel Julian date is local time and time_t is UTC.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <chrono>
#include <expected>
#include <limits>
#include <timezoneapi.h>

namespace xll {

	// UTC = local time + bias
	inline std::expected<LONG, DWORD> timezone_bias()
	{
		DYNAMIC_TIME_ZONE_INFORMATION dtzi;
		DWORD ret = GetDynamicTimeZoneInformation(&dtzi);
		if (TIME_ZONE_ID_INVALID == ret) {
			return std::unexpected<DWORD>(ret);
		}

		return (dtzi.Bias + dtzi.DaylightBias) * 60; // in seconds
	}

	// No timezone adjustment
	inline time_t excel_to_time_t(double jd)
	{
		return static_cast<time_t>((jd - 25569) * 86400);
	}
	inline double time_t_to_excel(time_t t)
	{
		return static_cast<double>(t) / 86400 + 25569;
	}

	// Convert Excel Julian local date to UTC time time_t
	inline time_t to_time_t(double jd)
	{
		// Excel Julian date is days since 1900-01-00
		// time_t is seconds since 1970-01-01
		// 70 years, 17 leap years, 1 day
		const auto bias = timezone_bias();
		if (!timezone_bias()) {
			return -1;
		}

		return static_cast<time_t>(excel_to_time_t(jd) + *bias);
	}

	// Convert UTC time time_t to Excel local Julian date
	inline double from_time_t(time_t t)
	{
		// Excel Julian date is days since 1900-01-00
		// time_t is seconds since 1970-01-01
		// 70 years, 17 leap years, 1 day
		const auto bias = timezone_bias();
		if (!bias) {
			return std::numeric_limits<double>::quiet_NaN();
		}

		return time_t_to_excel(t - *bias);
	}

	// Excel Julian date to sys_days
	inline std::chrono::sys_days to_days(double jd)
	{
		auto tp = std::chrono::system_clock::from_time_t(excel_to_time_t(jd));

		return std::chrono::time_point_cast<std::chrono::days>(tp);
	}
	inline double from_days(std::chrono::sys_days sd)
	{
		auto tp = std::chrono::time_point_cast<std::chrono::seconds>(sd);

		return time_t_to_excel(std::chrono::system_clock::to_time_t(tp));
	}

	// Excel Julian date to year_month_day
	inline std::chrono::year_month_day to_ymd(double jd)
	{
		return std::chrono::year_month_day{ to_days(jd) };
	}
	inline double from_ymd(std::chrono::year_month_day ymd)
	{
		return from_days(std::chrono::sys_days{ ymd });
	}

} // namespace xll