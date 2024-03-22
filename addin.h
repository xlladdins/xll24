// addin.h - Create Excel add-ins for xlAutoXXX functions.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <map>
#include "register.h"

namespace xll {

	class AddIn {
		Args args;

		// Auto<Register> function with Excel
		void Register()
		{
			if (addins.contains(args.functionText)) {
				const auto err = OPER(L"AddIn: ")
					& args.functionText & OPER(L" already registered");
				XLL_WARNING(view(err));
			}
			else {
				addins[args.functionText] = &args;
				const Auto<xll::Register> xao_reg([&]() -> int {
					try {
						OPER regId = XlfRegister(&args);

						return regId.xltype == xltypeNum;
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
		static inline std::map<OPER, Args*> addins;
		
		AddIn(const Args& args_)
			: args(args_)
		{
			Register();
			Unregister();
		}
		AddIn& operator=(const AddIn&) = delete;
		~AddIn()
		{ }
	};

} // namespace xll
