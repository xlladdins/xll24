// addin.h - Store data Excel needs to register add-ins.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <map>
#include "register.h"

namespace xll {

	inline std::map<OPER, Args*>& AddIns()
	{
		static std::map<OPER, Args*> addins;

		return addins;
	}
	inline std::map<double, Args*>& RegIds()
	{
		static std::map<double, Args*> regids;

		return regids;
	}

	// Find Args pointer given name or register id.
	inline const Args* find(const OPER& name)
	{
		const Args* pargs = nullptr;

		if (isNum(name)) {
			const auto i = RegIds().find(Num(name));
			if (i != RegIds().end()) {
				pargs = i->second;
			}
		}
		else {
			const auto i = AddIns().find(name);
			if (i != AddIns().end()) {
				pargs = i->second;
			}
		}

		return pargs;
	}

	// Create add-in to be registered with Excel.
	class AddIn {
		Args args;

		// Auto<Register> function with Excel
		void Register()
		{
			if (AddIns().contains(args.functionText)) {
				const auto err = OPER(L"AddIn: ")
					& args.functionText & OPER(L" already registered");
				XLL_WARNING(view(err));
			}
			AddIns()[args.functionText] = &args;
			const Auto<xll::Register> xao_reg([&]() -> int {
				try {
					OPER regid = XlfRegister(&args);
					if (regid.xltype == xltypeNum) {
						RegIds()[regid.val.num] = &args;
					}
					else {
						const auto err = OPER(L"AddIn: failed to register: ") & args.functionText;
						XLL_WARNING(view(err));
					}

					return regid.xltype == xltypeNum;
				}
				catch (const std::exception& ex) {
					XLL_ERROR(ex.what());

					return false;
				}
				catch (...) {
					XLL_ERROR("AddIn::Auto<Register>: unknown exception");

					return false;
				}
			});
		}

		// Auto<Unregister> function with Excel
		void Unregister()
		{
			const Auto<xll::Unregister> xao_unreg([text = args.functionText]() {
				try {
					if (!XlfUnregister(text)) {
						const auto err = OPER(L"AddIn: failed to unregister: ") & text;
						XLL_WARNING(view(err));

						return FALSE;
					}
					AddIns().erase(text);
					RegIds().erase(Num(Excel(xlfEvaluate, text)));
				}
				catch (const std::exception& ex) {
					XLL_ERROR(ex.what());

					return FALSE;
				}
				catch (...) {
					XLL_ERROR("AddIn::Auto<Unregister>: unknown exception");

					return FALSE;
				}

				return TRUE;
			});
		}
	public:
		// Lookup using function text or register id.
		static inline const Args* find(const OPER& text)
		{
			const Args* pargs = nullptr;

			if (isNum(text)) {
				const auto i = RegIds().find(Num(text));
				if (i != RegIds().end()) {
					pargs = i->second;
				}
			}
			else {
				const auto i = AddIns().find(text);
				if (i != AddIns().end()) {
					pargs = i->second;
				}
			}

			return pargs;
		}
		
		AddIn(const Args& args_)
			: args(args_)
		{
			Register();
			Unregister();
		}
		AddIn(const AddIn&) = delete;
		AddIn& operator=(const AddIn&) = delete;
		~AddIn()
		{ }
	};

} // namespace xll
